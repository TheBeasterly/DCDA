##Windows LibInstall
# Target Application Folder 
$APP_FOLDER = "OPPCAnalyzer"
if (-not (Test-Path -Path $APP_FOLDER)) {
    New-Item -ItemType Directory -Path $APP_FOLDER
}

# Python Download and Installation

# Function to fetch the latest Python version
function Get-LatestPythonVersion {
    $downloadPageUrl = "https://www.python.org/downloads/windows/"
    $webRequest = Invoke-WebRequest -Uri $downloadPageUrl
 
    # Basic HTML parsing (adjust if the download page structure changes significantly)
    $latestVersion = $webRequest.Content -match "Latest Python 3 Release - Python (\d+\.\d+\.\d+)"
    if ($matches) {
        return $matches[1]
    } else {
        Write-Error "Could not determine latest Python version."
        return $null 
    }
}

$latestVersion = Get-LatestPythonVersion
if (!$latestVersion) { Exit 1 } # Exit if we couldn't get the version

$downloadUrl = "https://www.python.org/ftp/python/$latestVersion/python-$latestVersion-amd64.exe" 
$destinationFile = ".\python-$latestVersion-amd64.exe"

# Check if Python is already installed (basic check)
if (!(Get-Command python -ErrorAction SilentlyContinue)) { 
    Write-Host "Downloading Python $latestVersion..."
    Invoke-WebRequest -Uri $downloadUrl -OutFile $destinationFile

    Write-Host "Installing Python..."
    Start-Process -FilePath $destinationFile -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait 
} else {
    Write-Host "Python is already installed"
}

# Virtual Environment Creation and Setup
$venvPath = $APP_FOLDER + "\oppca_venv"

Write-Host "Creating virtual environment..."
python -m venv $venvPath

Write-Host "Activating virtual environment..."
Set-Location $venvPath\Scripts
.\activate

# Package Installation
Write-Host "Upgrading pip..."
python -m pip install --upgrade pip

$packages = "tcl", "tk", "openpyxl", "pandas", "xlsxwriter"
foreach ($pkg in $packages) {
    if (Get-Module -ListAvailable -Name $pkg) {
        Write-Host "Upgrading $pkg..."
    } else {
        Write-Host "Installing $pkg..."
    }
    pip install --upgrade $pkg  
}