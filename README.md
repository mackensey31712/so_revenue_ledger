# Service Optimization Revenue Ledger

This repository contains scripts for analyzing and managing subscription data for the Service Optimization Revenue Dashboard.

## identify_multi_entry_accounts.py

This script analyzes Closed Won Opportunities data to identify accounts that have:
1. Multiple "Add Products" entries for the same product code (350-0100)
2. No "Reduction" or "Debook" entries after their latest "Add Products" entry
3. An "Active" account status

### Usage
```
python identify_multi_entry_accounts.py
```

The script generates a CSV report listing all accounts that meet these criteria, including:
- Account names
- Dates of all "Add Products" entries
- Number of "Add Products" entries
- Latest "Add Products" date

## Quick_Assist_Ledger_V2.py

This script processes closed won opportunities data and categorizes subscriptions based on their current status. It handles various edge cases including:

- Active subscriptions
- Ended subscriptions
- Churned accounts
- Inactive accounts
- Swapped accounts
- Free accounts

### Features

- Processes CSV files with opportunity data
- Categorizes accounts based on subscription status
- Handles special cases for specific accounts
- Generates detailed output with "Note" field indicating subscription status
- Archives processed files

### Requirements

- Python 3.x
- pandas
- datetime
- dateutil

### Usage

1. Place CSV files with opportunity data in the input directory
2. Run the script
3. Check the output directory for processed files

## License

Internal use only
