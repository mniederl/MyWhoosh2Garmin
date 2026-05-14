#!/usr/bin/env python3
"""
One-time Garmin browser login for MyWhoosh2Garmin.

Garmin currently blocks garth's direct username/password login for some users.
This opens a real browser, lets Garmin handle login/MFA, extracts the SSO
ticket from the redirect, exchanges it for garth tokens, and saves those tokens
to this repo's .garth directory.
"""
import re
from pathlib import Path
from urllib.parse import urlencode, urlparse, parse_qs

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

import garth
from garth.http import Client
from garth.sso import GarminOAuth1Session, OAUTH_USER_AGENT, exchange


SCRIPT_DIR = Path(__file__).resolve().parent
TOKENS_PATH = SCRIPT_DIR / ".garth"
GARMIN_BROWSER_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)

SSO_BASE_URL = "https://sso.garmin.com/sso"
SSO_EMBED_URL = f"{SSO_BASE_URL}/embed"
SIGNIN_URL = f"{SSO_BASE_URL}/signin?" + urlencode({
    "id": "gauth-widget",
    "embedWidget": "true",
    "gauthHost": SSO_BASE_URL,
    "service": SSO_EMBED_URL,
    "source": SSO_EMBED_URL,
    "redirectAfterAccountLoginUrl": SSO_EMBED_URL,
    "redirectAfterAccountCreationUrl": SSO_EMBED_URL,
})


def extract_ticket(url: str) -> str | None:
    parsed = urlparse(url)
    ticket = parse_qs(parsed.query).get("ticket", [None])[0]
    if ticket:
        return ticket
    match = re.search(r"[?&]ticket=(ST-[^&\s]+)", url)
    return match.group(1) if match else None


def configure_client(client: Client) -> None:
    client.sess.headers.update({"User-Agent": GARMIN_BROWSER_USER_AGENT})


def get_oauth1_token_from_browser_ticket(ticket: str, client: Client):
    sess = GarminOAuth1Session(parent=client.sess)
    base_url = f"https://connectapi.{client.domain}/oauth-service/oauth/"
    url = (
        f"{base_url}preauthorized?ticket={ticket}"
        f"&login-url={SSO_EMBED_URL}"
        "&accepts-mfa-tokens=true"
    )
    response = sess.get(
        url,
        headers={**OAUTH_USER_AGENT, "User-Agent": GARMIN_BROWSER_USER_AGENT},
        timeout=client.timeout,
    )
    response.raise_for_status()

    from urllib.parse import parse_qs
    from garth.auth_tokens import OAuth1Token

    parsed = parse_qs(response.text)
    token = {key: value[0] for key, value in parsed.items()}
    return OAuth1Token(domain=client.domain, **token)


def exchange_ticket_for_tokens(ticket: str) -> None:
    client = Client()
    configure_client(client)
    oauth1 = get_oauth1_token_from_browser_ticket(ticket, client)
    oauth2 = exchange(oauth1, client, login=True)
    client.configure(oauth1_token=oauth1, oauth2_token=oauth2)
    client.dump(str(TOKENS_PATH))

    print(f"Saved Garmin tokens to {TOKENS_PATH}")


def main() -> None:
    print("Opening Garmin login in Chromium.")
    print("Log in normally, including MFA if Garmin asks for it.")
    print("This script does not read or store your Garmin password.")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context(user_agent=GARMIN_BROWSER_USER_AGENT)
        page = context.new_page()
        page.goto(SIGNIN_URL, wait_until="domcontentloaded")

        try:
            page.wait_for_url(re.compile(r".*[?&]ticket=ST-[^&]+.*"), timeout=180_000)
            ticket_url = page.url
        except PlaywrightTimeoutError as error:
            raise RuntimeError(
                "Timed out waiting for Garmin's SSO ticket. "
                "Finish login/MFA in the opened browser window and rerun if needed."
            ) from error
        finally:
            browser.close()

    ticket = extract_ticket(ticket_url)
    if not ticket:
        raise RuntimeError(f"No Garmin SSO ticket found in URL: {ticket_url}")

    exchange_ticket_for_tokens(ticket)
    print("You can now run: python .\\myWhoosh2Garmin.py")


if __name__ == "__main__":
    main()
