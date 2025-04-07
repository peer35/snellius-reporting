# Snellius reporting by user and department
This script takes a Snellius xlsx report as input and will group usage by user (as opposed to by CPU/GPU/projectspace). It will also query the AD using ldap to find faculty and department information.

Usage will be calculated by year so the output can be used for calculating year-end costs.

## Installing
- Clone from github
- Create virtual environment and activate
- Install `openpyxl` and `ldap3`
```
pip install -r .\requirements.txt
```
- Sync the Teams folder "Add shortcut to Onedrive" 
- Copy `config.template.py` to `config.py` and adjust contents

## How to run
- Activate the venv 
- Run
```
python user_report.py
```
## Caching user data
User information found with the AD lookup will be stored in `data/userdata.json`. If an email address cannot be found the script will use the account information found in this file,
