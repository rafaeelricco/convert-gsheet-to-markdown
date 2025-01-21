<#
            _                      
   _____   (_)  _____  _____  ____ 
  / ___/  / /  / ___/ / ___/ / __ \
 / /     / /  / /__  / /__  / /_/ /
/_/     /_/   \___/  \___/  \____/ 
                                   
© r1cco.com

Environment Setup Module

This module is responsible for configuring the necessary Python virtual environment
to run the project scripts. It performs the following operations:

1. Verifies if Python is installed on the system
2. Creates a virtual environment if it doesn't exist
3. Activates the virtual environment in PowerShell
4. Installs the project dependencies via pip

The script is designed to work on Windows through PowerShell and manage
project dependencies via pip.
#>

$pythonCommand = if (Get-Command python -ErrorAction SilentlyContinue) { "python" } elseif (Get-Command python3 -ErrorAction SilentlyContinue) { "python3" } else { $null }

if (-not $pythonCommand) {
    Write-Host "Python not found. Install Python first."
    exit 1
}

$venvDir = "env"
$activateScript = Join-Path $venvDir "Scripts" "Activate.ps1"

if (-not (Test-Path $venvDir)) {
    Write-Host "Creating virtual environment..."
    & $pythonCommand -m venv $venvDir
}

if (Test-Path $activateScript) {
    Write-Host "Activating virtual environment..."
    . $activateScript
} else {
    Write-Host "Error: Activation script not found in $activateScript"
    exit 1
}

if (Test-Path "requirements.txt") {
    Write-Host "Installing dependencies..."
    pip install -r requirements.txt
} else {
    Write-Host "Warning: requirements.txt not found"
}
    
Write-Host "`nEnvironment ready! To reactivate manually use:"
Write-Host "  .\$venvDir\Scripts\Activate.ps1`n"