#!/bin/bash

"""
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
3. Activates the virtual environment across multiple platforms (Linux, macOS, Windows)
4. Installs the project dependencies via pip

The script is designed to work across different operating systems and manage
project dependencies via pip.
"""

if ! command -v python3 &> /dev/null && ! command -v python &> /dev/null; then
    echo "Python not found. Install Python first."
    exit 1
fi

VENV_DIR="env"
if [ ! -d "$VENV_DIR" ]; then
    if command -v python3 &> /dev/null; then
        python3 -m venv "$VENV_DIR"
    else
        python -m venv "$VENV_DIR"
    fi
fi

case "$(uname -s)" in
    Linux*|Darwin*)
        source "$VENV_DIR/bin/activate"
        ;;
    MINGW*|CYGWIN*|MSYS*)
        source "$VENV_DIR/Scripts/activate"
        ;;
    *)
        echo "System not recognized. For Windows, run in PowerShell:"
        echo ".\$VENV_DIR\Scripts\Activate.ps1"
        exit 1
        ;;
esac

pip install -r requirements.txt

echo "Environment setup completed successfully!"