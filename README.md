# myWhoosh2Garmin

## Features

* Finds the `.fit` files from your MyWhoosh installation.
* Fixes missing power and heart rate averages.
* Removes temperature.
* Creates a backup file to a folder you select.
* Uploads the fixed `.fit` file to Garmin Connect.
* Renames uploaded Garmin activities from MyWhoosh custom workout metadata when available.

## Installation Steps

1. Download `myWhoosh2Garmin.py` and `garmin_browser_login.py` to your filesystem to a folder of your choosing.

2. Go to the folder where you downloaded the scripts in a shell.

* **MacOS:** Terminal of your choice.
* **Windows:** Start > Run > cmd or Start > Run > powershell

3. Install `pipenv` if not already installed:

```shell
pip3 install pipenv
or
pip install pipenv
```

4. Install dependencies in a virtual environment:

```shell
pipenv install
```

5. Activate the virtual environment:

```shell
pipenv shell
```

6. Run the Garmin browser login once:

```shell
python3 garmin_browser_login.py
or
python garmin_browser_login.py
```

This opens a Chromium browser window. Log in to Garmin Connect normally, including MFA if prompted. The script saves reusable Garmin tokens to the local `.garth` directory, because direct username/password login through Garth no longer reliably works with Garmin's current auth flow.

7. Choose your backup folder.

### MacOS

![image](https://github.com/user-attachments/assets/2c6c1072-bacf-4f0c-8861-78f62bf51648)

### Windows

![image](https://github.com/user-attachments/assets/d1540291-4e6d-488e-9dcf-8d7b68651103)

8. Run the script when you're done riding or running.

```shell
python3 myWhoosh2Garmin.py
or
python myWhoosh2Garmin.py
```

Optional: use Zwift-like FIT device metadata before upload:

```shell
python3 myWhoosh2Garmin.py --fix-device
or
python myWhoosh2Garmin.py --fix-device
```

This changes the exported FIT creator metadata to look more like a Zwift virtual ride by setting the FIT manufacturer to Zwift and adding creator device information.

After uploading, the script tries to rename the Garmin activity automatically. It reads the MyWhoosh session UUID from the FIT file, loads the matching cached custom workout JSON, and renames the Garmin activity to `MyWhoosh - <workout name>`. It only updates an activity when the Garmin activity matches the FIT start time, duration, and distance.

Example output:

```text
2024-11-21 10:08:37,107 Checking for .fit files in directory: <YOUR_MYWHOOSH_DIR_WITH_FITFILES>.
2024-11-21 10:08:37,107 Found the most recent .fit file: MyNewActivity-3.8.5.fit.
2024-11-21 10:08:37,107 Cleaning up <YOUR_BACKUP_FOLDER>yNewActivity-3.8.5_2024-11-21_100837.fit.
2024-11-21 10:08:37,855 Cleaned-up file saved as <YOUR_BACKUP_FOLDER>MyNewActivity-3.8.5_2024-11-21_100837.fit
2024-11-21 10:08:37,871 Successfully cleaned MyNewActivity-3.8.5.fit and saved it as MyNewActivity-3.8.5_2024-11-21_100837.fit.
2024-11-21 10:08:38,408 Duplicate activity found on Garmin Connect.
```

9. Or see below to automate the process.

## Automation Tips

What if you want to automate the whole process:

### MacOS

PowerShell on MacOS (Verified & works)

You need Powershell

```shell
brew install powershell/tap/powershell
```

```powershell
# Define the JSON config file path
$configFile = "$PSScriptRoot\mywhoosh_config.json"
$myWhooshApp = "MyWhoosh Indoor Cycling App.app"

# Check if the JSON file exists and read the stored path
if (Test-Path $configFile) {
    $config = Get-Content -Path $configFile | ConvertFrom-Json
    $mywhooshPath = $config.path
} else {
    $mywhooshPath = $null
}

# Validate the stored path
if (-not $mywhooshPath -or -not (Test-Path $mywhooshPath)) {
    Write-Host "Searching for $myWhooshApp"
    $mywhooshPath = Get-ChildItem -Path "/Applications" -Filter $myWhooshApp -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

    if (-not $mywhooshPath) {
        Write-Host " not found!"
        exit 1
    }

    $mywhooshPath = $mywhooshPath.FullName

    # Store the path in the JSON file
    $config = @{ path = $mywhooshPath }
    $config | ConvertTo-Json | Set-Content -Path $configFile
}

Write-Host "Found $myWhooshApp at $mywhooshPath"

Start-Process -FilePath $mywhooshPath

# Wait for the application to finish
Write-Host "Waiting for $myWhooshApp to finish..."
while ($process = ps -ax | grep -i $myWhooshApp | grep -v "grep") {
    Write-Output $process
    Start-Sleep -Seconds 5
}

# Run the Python script
Write-Host "$myWhooshApp has finished, running Python script..."
python3 "<PATH_WHERE_YOUR_SCRIPT_IS_LOCATED>/MyWhoosh2Garmin/myWhoosh2Garmin.py"
```

AppleScript (need to test further)

```applescript
TODO: needs more work
```

### Windows

Windows `.ps1` (PowerShell) file (Untested on Windows)

```powershell
# Define the JSON config file path
$configFile = "$PSScriptRoot\mywhoosh_config.json"

# Check if the JSON file exists and read the stored path
if (Test-Path $configFile) {
    $config = Get-Content -Path $configFile | ConvertFrom-Json
    $mywhooshPath = $config.path
} else {
    $mywhooshPath = $null
}

# Validate the stored path
if (-not $mywhooshPath -or -not (Test-Path $mywhooshPath)) {
    Write-Host "Searching for mywhoosh.exe..."
    $mywhooshPath = Get-ChildItem -Path "C:\PROGRAM FILES" -Filter "mywhoosh.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

    if (-not $mywhooshPath) {
        Write-Host "mywhoosh.exe not found!"
        exit 1
    }

    $mywhooshPath = $mywhooshPath.FullName

    # Store the path in the JSON file
    $config = @{ path = $mywhooshPath }
    $config | ConvertTo-Json | Set-Content -Path $configFile
}

Write-Host "Found mywhoosh.exe at $mywhooshPath"

# Start mywhoosh.exe
Start-Process -FilePath $mywhooshPath

# Wait for the application to finish
Write-Host "Waiting for mywhoosh to finish..."
while (Get-Process -Name "mywhoosh" -ErrorAction SilentlyContinue) {
    Start-Sleep -Seconds 5
}

# Run the Python script
Write-Host "mywhoosh has finished, running Python script..."
python "C:\Path\to\myWhoosh2Garmin.py"
```

## Built With

Technologies used in the project:

* Neovim
* [Garth](https://github.com/matin/garth)
* tKinter
* [Fit_tool](https://bitbucket.org/stagescycling/python_fit_tool/src/main/)
