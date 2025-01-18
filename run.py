"""
            _                      
   _____   (_)  _____  _____  ____ 
  / ___/  / /  / ___/ / ___/ / __ \
 / /     / /  / /__  / /__  / /_/ /
/_/     /_/   \___/  \___/  \____/ 
                                   
© r1cco.com

Script Runner Module

This module provides functionality to sequentially execute Python scripts while handling errors
and providing execution status feedback. It specifically manages the execution of two scripts:
1. access_gsheet_and_save_data.py
2. convert-sheets-to-markdown.py

The scripts are executed in order, and the process stops if the first script fails.
"""

import os
import subprocess
import sys


def run_script(script_path):
    """
    Execute a Python script and verify its execution status.

    Args:
        script_path (str): The path to the Python script to be executed.

    Returns:
        bool: True if the script executed successfully (return code 0),
              False if there was an error during execution.

    Raises:
        subprocess.CalledProcessError: If the script execution fails.
    """
    print(f"\nExecutando {os.path.basename(script_path)}...")
    try:
        result = subprocess.run([sys.executable, script_path], check=True)
        if result.returncode == 0:
            print(f"✓ {os.path.basename(script_path)} executado com sucesso!")
            return True
        return False
    except subprocess.CalledProcessError as e:
        print(f"✗ Erro ao executar {os.path.basename(script_path)}: {e}")
        return False


def main():
    """
    Main function that orchestrates the script execution process.

    This function:
    1. Defines the paths for the required scripts
    2. Verifies that all scripts exist
    3. Executes the scripts in sequence
    4. Handles execution failures by terminating the process if needed

    Returns:
        None

    Exits:
        1: If any script file is not found or if the first script fails
    """
    script1 = os.path.join("scripts", "access_gsheet_and_save_data.py")
    script2 = os.path.join("scripts", "convert-sheets-to-markdown.py")

    for script in [script1, script2]:
        if not os.path.exists(script):
            print(f"Erro: O arquivo {script} não foi encontrado.")
            sys.exit(1)

    if run_script(script1):
        run_script(script2)
    else:
        print("\nProcesso interrompido devido a erro no primeiro script.")
        sys.exit(1)


if __name__ == "__main__":
    main()
