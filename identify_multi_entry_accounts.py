import pandas as pd
import glob
import os
from datetime import datetime

def identify_multi_entry_accounts():
    """
    Identifies accounts that have:
    1. Multiple "Add Products" entries for the same product code
    2. No "Reduction" or "Debook" entries after their latest "Add Products" entry
    3. An "Active" account status
    """
    # Define directory path for input files
    input_dir = r"C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Test Files"
    
    # Find all CSV files with "FQA Closed Won Opportunities" in the name
    csv_files = glob.glob(os.path.join(input_dir, "FQA Closed Won Opportunities*.csv"))
    
    if not csv_files:
        print(f"No Closed Won Opportunities CSV files found in {input_dir}")
        # Try with a more generic pattern
        csv_files = glob.glob(os.path.join(input_dir, "*Opportunities*.csv"))
        if not csv_files:
            print("No Opportunities CSV files found either. Please check file names.")
            return
    
    # Use the most recent file based on filename
    most_recent_file = max(csv_files)
    print(f"Using file: {most_recent_file}")
    
    # Load the CSV file
    try:
        df = pd.read_csv(most_recent_file)
    except Exception as e:
        print(f"Error loading file: {e}")
        return
    
    # Check if required columns exist
    required_cols = ['Account Name', 'Close Date', 'Product Code', 
                    'Opportunity Type', 'Account Status']
    
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        print(f"Missing columns: {missing_cols}")
        print("Available columns:", df.columns.tolist())
        return
    
    # Convert 'Close Date' to datetime format
    df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
    
    # Filter for Service Optimization entries (Product Code 350-0100)
    so_df = df[df['Product Code'] == '350-0100'].copy()
    
    # Sort by Account Name and Close Date
    so_df = so_df.sort_values(by=['Account Name', 'Close Date']).reset_index(drop=True)
    
    # Initialize a list to store results
    multi_entry_accounts = []
    
    # Get unique account names
    account_names = so_df['Account Name'].unique()
    
    # Process each account
    for account_name in account_names:
        # Get all entries for this account
        account_entries = so_df[so_df['Account Name'] == account_name].copy()
        
        # Check if account has multiple "Add Products" entries
        add_products_entries = account_entries[account_entries['Opportunity Type'] == 'Add Products']
        if len(add_products_entries) < 2:
            continue
        
        # Check if account is Active
        latest_entry = account_entries.iloc[-1]
        if latest_entry['Account Status'] != 'Active':
            continue
        
        # Check if there are no Reduction or Debook entries after the latest Add Products entry
        latest_add_date = add_products_entries['Close Date'].max()
        
        reduction_after_latest_add = account_entries[
            (account_entries['Close Date'] > latest_add_date) & 
            (account_entries['Opportunity Type'].isin(['Reduction', 'Debook']))
        ]
        
        if len(reduction_after_latest_add) > 0:
            continue
            
        # If we got here, the account meets all criteria
        add_products_dates = add_products_entries['Close Date'].dt.strftime('%m/%d/%Y').tolist()
        multi_entry_accounts.append({
            'Account Name': account_name,
            'Add Products Dates': add_products_dates,
            'Number of Add Products': len(add_products_entries),
            'Latest Add Date': latest_add_date.strftime('%m/%d/%Y'),
            'Account Status': latest_entry['Account Status']
        })
    
    # Create a DataFrame from the results
    if multi_entry_accounts:
        results_df = pd.DataFrame(multi_entry_accounts)
        
        # Generate timestamp for the output file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'Multi_Entry_Accounts_{timestamp}.csv'
        output_path = os.path.join(input_dir, output_file)
        
        # Save results to CSV
        results_df.to_csv(output_path, index=False)
        print(f"\nFound {len(multi_entry_accounts)} accounts that meet the criteria.")
        print(f"Results saved to: {output_path}")
        
        # Print the list of accounts
        print("\nList of multi-entry active accounts:")
        for idx, account in enumerate(multi_entry_accounts, 1):
            print(f"{idx}. {account['Account Name']} - {account['Number of Add Products']} Add Products entries")
            print(f"   Latest Add Date: {account['Latest Add Date']}")
            print(f"   All Add Dates: {', '.join(account['Add Products Dates'])}")
    else:
        print("\nNo accounts found that meet all criteria.")

if __name__ == "__main__":
    identify_multi_entry_accounts()
