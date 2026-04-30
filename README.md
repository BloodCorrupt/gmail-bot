# Gmail Account Checker / Updater

Python automation script that signs in to Gmail accounts from an Excel sheet and optionally updates account fields (recovery email, password, first name, last name) using Selenium.

## What this project does

- Loads account rows from `account.xlsx`
- Attempts Gmail login for each row
- Handles common sign-in branches (password, recovery prompt, some account-state checks)
- Optionally applies profile/security updates based on Excel values
- Writes status/result back to column 8 in `account.xlsx`
- Runs accounts concurrently using a thread pool

## Requirements

- Windows (tested in this workspace)
- Python 3.12+
- Google Chrome installed
- Python packages listed in `requirements.txt`

Selenium 4 uses Selenium Manager, so a separate manual ChromeDriver install is usually not required.

## Installation

From the project folder, install dependencies:

```powershell
C:/Users/<your username>/AppData/Local/Programs/Python/Python312/python.exe -m pip install -r requirements.txt
```

## Input file format (`account.xlsx`)

The script expects these columns in the first sheet:

1. `email` (required)
2. `password` (required)
3. `recovery_email` (optional; used if recovery verification appears)
4. `new_recovery_email` (optional; if present, script tries to update recovery email)
5. `new_password` (optional; if present, script tries to update password)
6. `new_first_name` (optional; if present, script tries to update first name)
7. `new_last_name` (optional; if present, script tries to update last name)
8. `status` (output column written by the script)

Only rows with both `email` and `password` are processed.

## Run

```powershell
C:/Users/<your username>/AppData/Local/Programs/Python/Python312/python.exe gmail-checker.py
```

You will be prompted for thread count.

## Common issues

- `No module named ...`
  - Reinstall dependencies with the same Python executable used to run the script.
- `No accounts found in account.xlsx`
  - Ensure `account.xlsx` exists in the project folder and has data rows starting from row 2.
- Browser/session instability
  - Lower thread count (for example, `1-3`) and retry.

## Notes

- This script automates account actions that may trigger Google security checks.
- Account security workflows can change at any time and may break selectors/steps.
- Keep sensitive account files out of source control.

## Project files

- `gmail-checker.py` - main script
- `requirements.txt` - Python dependencies
- `.gitignore` - ignore rules
- `account.xlsx` - local input/output workbook
