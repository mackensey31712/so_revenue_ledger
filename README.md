# Service Optimization Revenue Ledger

This repository contains scripts for analyzing and managing subscription data for the Service Optimization Revenue Dashboard.

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
