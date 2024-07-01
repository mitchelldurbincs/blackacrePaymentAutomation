# Check if Git is installed
$gitCommand = Get-Command git -ErrorAction SilentlyContinue
if (-not $gitCommand) {
    Write-Host "Git not found. Installing Git..."
    # Download Git installer
    Invoke-WebRequest -Uri "https://github.com/git-for-windows/git/releases/download/v2.33.0.windows.2/Git-2.33.0.2-64-bit.exe" -OutFile "git_installer.exe"
    # Install Git
    Start-Process -FilePath "git_installer.exe" -ArgumentList "/VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS /COMPONENTS='icons,ext\reg\shellhere,assoc,assoc_sh'" -Wait
    # Refresh environment variables
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
	Write-Host "Finished installing Git..."
}

# Check if Python is installed
$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCommand) {
    Write-Host "Python not found. Installing Python..."
	# Download Python installer
	Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.11.0/python-3.11.0-amd64.exe" -OutFile "python_installer.exe"
	# Install Python (adjust the version number if needed)
	Start-Process -FilePath "python_installer.exe" -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
	# Refresh environment variables
	$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
	Write-Host "Finished installing Python..."
} else {
    $pythonVersion = (python --version 2>&1).ToString().Split()[1]
    Write-Host "Python version $pythonVersion is already installed."
}

# Clone the repository
Write-Host "Cloning the repository..."
git clone https://github.com/mitchelldurbincs/blackacrePaymentAutomation
Set-Location -Path "blackacrePaymentAutomation"

# Install dependencies
Write-Host "Installing dependencies..."
pip install -r requirements.txt

Write-Host "Installation complete! You can now run the program."