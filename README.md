# Quick Assist Ledger Processing Script

## Overview

The `Quick_Assist_Ledger_V4.py` script is a comprehensive data processing tool designed to transform Five9 Closed Won Opportunities data into a structured ledger format suitable for revenue tracking and BigQuery analysis. The script processes subscription-based service optimization accounts and generates intermediate billing entries to create a complete monthly revenue trail.

## Features

### Core Functionality
- **Automated CSV Processing**: Processes all CSV files in the input directory
- **Subscription Lifecycle Management**: Handles start, intermediate billing, partial reductions, and end of subscriptions
- **Multiple Subscription Support**: Manages accounts with multiple concurrent subscriptions
- **Special Case Handling**: Includes custom logic for specific accounts and scenarios
- **BigQuery Compatibility**: Formats output with underscore-separated column names
- **File Management**: Automatically archives processed files and copies scripts to production

### Data Processing Capabilities
1. **Duplicate Detection**: Identifies and marks duplicate records
2. **Intermediate Entry Generation**: Creates monthly billing entries between start and end dates
3. **Subscription Status Tracking**: Categorizes accounts as active, ended, or other status
4. **Amount Calculations**: Handles partial reductions and amount adjustments
5. **Note Generation**: Adds descriptive notes for different transaction types

## Directory Structure

```
C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Revenue Dashboard\QuickAssist_Ledger\
├── In\                    # Input CSV files
├── Out\                   # Processed output files
├── Archived\              # Archived source files
└── Quick_Assist_Ledger_V4.py  # Production script copy
```

## Input Data Requirements

### Required Columns
- `Account Name`: Customer account identifier
- `Five9 Account Number`: Five9 specific account number
- `Close Date`: Transaction date
- `Product Code`: Product identifier (350-0100 for Service Optimization, 350-0101 for On Demand)
- `Amount`: Transaction amount
- `Opportunity Type`: Type of transaction (Add Products, Reduction, Debook)
- `Opportunity ID`: Unique opportunity identifier
- `Account ID`: Account identifier
- `Account Status`: Current account status (Active, Churned, etc.)

### Product Codes
- **350-0100**: Service Optimization (Subscription-based)
- **350-0101**: On Demand (Transaction-based)

## Processing Logic

### Service Optimization (350-0100)

#### Add Products (Positive Amount)
- **Start of Subscription**: Creates initial subscription entry
- **Intermediate Billing**: Generates monthly entries from start date to current month (for active single-entry accounts)
- **Multiple Subscriptions**: Handles accounts with multiple concurrent subscriptions using suffix notation (_2, _3, etc.)

#### Add Products (Negative Amount)
- **Partial Reduction**: Maintains subscription with reduced billing amount
- **Note**: "Reduction - Subscribed Billing"

#### Reduction Entries
- **Partial Reduction**: When reduction amount < original amount
- **Full Reduction**: When reduction amount >= original amount
- **Note**: "Reduction - End of Subscription" for full reductions

#### Debook Entries
- **End of Subscription**: Marks the termination of subscription
- **Note**: "Reduction - End of Subscription"

### On Demand (350-0101)
- **Add Products**: Marked as "On Demand Entry"
- **Reduction/Debook**: Marked as "Reduction"

### Special Account Processing

The script includes custom logic for specific accounts that require special intermediate date entries:

#### 1. ADT Solar LLC (fka SUNPRO)
- **Requirement**: Monthly entries from 2/2/2024 to 8/2/2024
- **Note**: "Subscribed Billing"

#### 2. Electronic Caregiver
- **Requirements**: 
  - Monthly entries from 4/21/2025 to 6/21/2025
  - Special entry on 7/14/2025 with amount $12.58
- **Note**: "Subscribed Billing"

#### 3. Sun Source Energy
- **Requirements**:
  - Special entry for 12/11/2024
  - Monthly entries from 1/11/2025 to 6/11/2025 with amount $414.75
- **Note**: "Subscribed Billing"

### Account Status Classification

#### Active Subscribers
- **Criteria**: 
  - Opportunity Type = "Add Products"
  - Note contains "Start of Subscription" or "Subscribed Billing"
  - Note does not contain "Churned", "Inactive", or "Swap"
  - Account Status = "Active"

#### Ended Subscriptions
- **Criteria**:
  - Notes containing "Reduction - End of Subscription"
  - Notes containing "Start of Subscription - Churned"
  - Notes containing "Start of Subscription - Inactive"
  - Notes containing "Start of Subscription - Swap"

#### Special Cases
- **Churned Accounts**: Account Status = "Churned" → "Start of Subscription - Churned"
- **Inactive Accounts**: Account Status = NaN/null → "Start of Subscription - Inactive"
- **Swap Accounts**: Amount = 0 with corresponding On Demand entry → "Start of Subscription - Swap"

## Output Format

### Column Structure (BigQuery Compatible)
- `Account_Name`: Customer account name
- `Five9_Account_Number`: Five9 account identifier
- `Date`: Transaction date (MM/DD/YYYY)
- `Product_Code`: Product code
- `Amount`: Transaction amount
- `Opportunity_Type`: Transaction type
- `Opportunity_ID`: Unique identifier
- `Account_ID`: Account identifier
- `Account_Status`: Account status
- `Note`: Descriptive transaction note

### File Naming Convention
`Quick_Assist_Ledger_Output_YYYYMMDD_HHMMSS.csv`

## Usage Instructions

### Prerequisites
- Python 3.7+
- Required packages: `pandas`, `python-dateutil`

### Installation
```bash
pip install pandas python-dateutil
```

### Running the Script
1. Place CSV files in the input directory:
   ```
   C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Revenue Dashboard\QuickAssist_Ledger\In\
   ```

2. Execute the script:
   ```bash
   python Quick_Assist_Ledger_V4.py
   ```

3. Check output in:
   ```
   C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Revenue Dashboard\QuickAssist_Ledger\Out\
   ```

### Automated Process Flow
1. **Input Detection**: Scans input directory for CSV files
2. **Data Validation**: Checks for required columns
3. **Processing**: Applies business logic and transformations
4. **Special Entries**: Adds custom intermediate entries for specified accounts
5. **Output Generation**: Creates timestamped output file
6. **Statistics Display**: Shows processing summary in terminal
7. **File Management**: Archives source files and copies script to production

## Output Statistics

The script provides comprehensive statistics including:
- Number of active subscribers
- List of active subscriber accounts
- Number of accounts with ended subscriptions
- List of ended subscription accounts
- Accounts with multiple subscriptions
- Total account counts across all categories

## Error Handling

- **Missing Columns**: Reports missing required columns and skips file
- **File Access Errors**: Handles file read/write permission issues
- **Data Type Errors**: Manages date parsing and numeric conversion errors
- **Directory Creation**: Automatically creates missing directories

## Maintenance and Updates

### Version History
- **V4**: Current version with special intermediate entries and enhanced multiple subscription handling
- **V3**: Added multiple subscription support
- **V2**: Enhanced intermediate entry generation
- **V1**: Basic processing functionality

### Future Enhancements
- Configuration file support for special account rules
- Email notifications for processing completion
- Database integration capabilities
- Enhanced error logging and monitoring

## Technical Notes

### Performance Considerations
- Processes files sequentially for data integrity
- Memory-efficient DataFrame operations
- Optimized date calculations using `relativedelta`

### Data Integrity
- Maintains referential integrity between related entries
- Preserves original transaction amounts and dates
- Implements consistent note formatting across all entries

### Compatibility
- Compatible with Python 3.7+
- Tested with pandas 1.x and 2.x
- Windows-optimized file path handling

## Support and Troubleshooting

### Common Issues
1. **No CSV files found**: Ensure files are in the correct input directory
2. **Missing columns error**: Verify input file contains all required columns
3. **Permission errors**: Check file/directory permissions
4. **Memory errors**: Process smaller batches of data

### Debug Mode
The script provides detailed console output showing:
- Files being processed
- Number of records processed
- Accounts with special processing
- Final statistics and file locations

## Legacy Scripts

This repository also contains legacy scripts for reference:

### identify_multi_entry_accounts.py
Analyzes Closed Won Opportunities data to identify accounts with multiple "Add Products" entries.

### Quick_Assist_Ledger_V2.py
Previous version of the ledger processing script with basic functionality.

## Contact Information

For questions, issues, or enhancement requests, please contact the development team or create an issue in the repository.

1. Place CSV files with opportunity data in the input directory
2. Run the script
3. Check the output directory for processed files

## License

Internal use only
