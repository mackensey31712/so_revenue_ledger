import pandas as pd
import glob
import os
from datetime import datetime
import calendar
from dateutil.relativedelta import relativedelta

def identify_multiple_subscription_accounts():
    """
    Identifies accounts that have:
    1. Multiple "Start of Subscription" entries
    2. No "Reduction" or "Debook" entries between ANY "Start of Subscription" entries
    3. An "Active" account status
    
    This helps identify accounts that might be double-billed or have overlapping subscriptions.
    """
    # Define directory path for input files
    input_dir = r"C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Test Files"
    
    # Look for Quick Assist Ledger output files first (preferred)
    ledger_files = glob.glob(os.path.join(input_dir, "Quick_Assist_Ledger_Output*.csv"))
    
    if not ledger_files:
        print("No Quick Assist Ledger output files found. Looking for raw opportunity files...")
        # Fall back to Opportunities files
        ledger_files = glob.glob(os.path.join(input_dir, "*Opportunities*.csv"))
        if not ledger_files:
            print("No suitable CSV files found. Please run Quick_Assist_Ledger_V2.py first or check file names.")
            return
    
    # Use the most recent file based on filename
    most_recent_file = max(ledger_files)
    print(f"Using file: {most_recent_file}")
    
    # Load the CSV file
    try:
        df = pd.read_csv(most_recent_file)
    except Exception as e:
        print(f"Error loading file: {e}")
        return
    
    # Check for required columns - adjusted for both original and processed files
    if 'Note' in df.columns:
        # This is a processed file from Quick_Assist_Ledger output
        note_col = 'Note'
        account_col = 'Account_Name' if 'Account_Name' in df.columns else 'Account Name'
        date_col = 'Date'
        product_col = 'Product_Code' if 'Product_Code' in df.columns else 'Product Code'
        status_col = 'Account_Status' if 'Account_Status' in df.columns else 'Account Status'
        opportunity_type_col = 'Opportunity_Type' if 'Opportunity_Type' in df.columns else 'Opportunity Type'
        amount_col = 'Amount'
    else:
        # This is a raw file, we can't proceed
        print("This file doesn't have a 'Note' column. Please run Quick_Assist_Ledger_V2.py first.")
        return
    
    # Convert date to datetime format
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    
    # Filter for Service Optimization entries (Product Code 350-0100)
    so_df = df[df[product_col] == '350-0100'].copy()
    
    # Sort by Account Name and Date
    so_df = so_df.sort_values(by=[account_col, date_col]).reset_index(drop=True)
    
    # Initialize a list to store results
    problematic_accounts = []
    
    # Get unique account names
    account_names = so_df[account_col].unique()
    
    # Process each account
    for account_name in account_names:
        # Get all entries for this account
        account_entries = so_df[so_df[account_col] == account_name].copy()
        
        # Skip if account is not active
        latest_status = account_entries[status_col].iloc[-1]
        if latest_status != 'Active':
            continue
        
        # Check for Start of Subscription entries
        start_sub_entries = account_entries[account_entries[note_col] == 'Start of Subscription']
        
        if len(start_sub_entries) < 2:
            continue
            
        # Sort the Start of Subscription entries by date
        start_sub_entries = start_sub_entries.sort_values(by=date_col)
        start_dates = start_sub_entries[date_col].tolist()
        
        # Check for any Reduction or Debook entries between ANY pair of Start of Subscription dates
        has_reduction_between = False
        
        for i in range(len(start_dates) - 1):
            current_date = start_dates[i]
            next_date = start_dates[i + 1]
            
            # Find entries between these two Start of Subscription dates
            entries_between = account_entries[
                (account_entries[date_col] > current_date) & 
                (account_entries[date_col] < next_date)
            ]
            
            # Check if any of these entries are Reductions or Deductions
            reduction_entries = entries_between[
                (entries_between[opportunity_type_col].isin(['Reduction', 'Debook'])) | 
                (entries_between[note_col].str.contains('Reduction', na=False))
            ]
            
            if len(reduction_entries) > 0:
                has_reduction_between = True
                break
        
        # If there are no reductions between any Start of Subscription entries, this is what we're looking for
        if not has_reduction_between:
            # Format the dates and amounts for output
            formatted_dates = [date.strftime('%m/%d/%Y') for date in start_dates]
            amounts = start_sub_entries[amount_col].tolist()
            
            # Add to our results
            problematic_accounts.append({
                'Account Name': account_name,
                'Account Status': latest_status,
                'Start of Subscription Dates': formatted_dates,
                'Amounts': amounts,
                'Count': len(start_sub_entries)
            })
    
    # Create a report with the results
    if problematic_accounts:
        # Generate timestamp for the output file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Create a detailed CSV
        rows = []
        for account in problematic_accounts:
            rows.append({
                'Account Name': account['Account Name'],
                'Account Status': account['Account Status'],
                'Start of Subscription Dates': ", ".join(account['Start of Subscription Dates']),
                'Amounts': ", ".join([str(amt) for amt in account['Amounts']]),
                'Count': account['Count']
            })
        
        if rows:
            results_df = pd.DataFrame(rows)
            output_file = f'Multiple_Subscription_Accounts_{timestamp}.csv'
            output_path = os.path.join(input_dir, output_file)
            results_df.to_csv(output_path, index=False)
            print(f"\nFound {len(problematic_accounts)} accounts with multiple subscriptions and no reductions between.")
            print(f"Results saved to: {output_path}")
            
            # Print the list of accounts
            print("\nList of accounts with multiple subscriptions:")
            for idx, account in enumerate(problematic_accounts, 1):
                print(f"{idx}. {account['Account Name']} - {account['Count']} Start of Subscription entries")
                for i in range(len(account['Start of Subscription Dates'])):
                    print(f"   - {account['Start of Subscription Dates'][i]}: ${account['Amounts'][i]}")
                print()
    else:
        print("\nNo accounts found with multiple subscriptions and no reductions between.")

if __name__ == "__main__":
    identify_multiple_subscription_accounts()
