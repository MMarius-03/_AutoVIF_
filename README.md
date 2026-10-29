# AutoVIF

Python desktop app for reading OF items from a CSV and entering them through a configured PuTTY session. The script's command line currently identifies the workflow as AutoOF.

## Requirements

- Windows
- Python 3
- PuTTY with the expected session configured
- The CSV and filter files set in `config.json`

## Run

Install the dependencies and check the CSV without opening PuTTY:

```powershell
py -m pip install -r requirements.txt
py putty_auto.py --dry-run
```

Run the automation with the configured PuTTY session:

```powershell
py putty_auto.py
```

Run the built-in checks with:

```powershell
py putty_auto.py --self-test
```
