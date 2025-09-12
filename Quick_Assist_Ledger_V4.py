import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
import os
from datetime import datetime
import glob
import shutil

def process_closed_won_opportunities():
    # Define directory paths
    input_dir = r"C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Revenue Dashboard\QuickAssist_Ledger\In"
    output_dir = r"C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Revenue Dashboard\QuickAssist_Ledger\Out"
    archive_dir = r"C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Revenue Dashboard\QuickAssist_Ledger\Archived"
    
    # Check if directories exist, create if they don't
    for dir_path in [input_dir, output_dir, archive_dir]:
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
            print(f"Created directory: {dir_path}")
    
    # Find all CSV files in the input directory
    csv_files = glob.glob(os.path.join(input_dir, "*.csv"))
    
    if not csv_files:
        print(f"No CSV files found in {input_dir}")
        return
    
    # Process each CSV file found
    for file_path in csv_files:
        process_file(file_path, output_dir, archive_dir)

def process_file(file_path, output_dir, archive_dir):
    file_name = os.path.basename(file_path)
    print(f"Processing file: {file_name}...")
    
    # Load the CSV file
    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        print(f"Error loading file {file_name}: {e}")
        return
    
    # Select only the required columns
    cols_to_keep = [
        'Account Name', 'Five9 Account Number', 'Close Date', 'Product Code', 
        'Amount', 'Opportunity Type', 'Opportunity ID', 'Account ID', 'Account Status'
    ]
    
    # Check if all required columns exist in the dataframe
    missing_cols = [col for col in cols_to_keep if col not in df.columns]
    if missing_cols:
        print(f"Missing columns in {file_name}: {missing_cols}")
        return
    
    # Filter out only the required columns
    df = df[cols_to_keep].copy()
    
    # Create a new 'Note' column
    df['Note'] = ""
    
    # Rename 'Close Date' to 'Date'
    df = df.rename(columns={'Close Date': 'Date'})
    
    # Convert 'Date' to datetime format
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    # Sort by 'Account Name', 'Product Code', then 'Date'
    df = df.sort_values(by=['Account Name', 'Product Code', 'Date']).reset_index(drop=True)
    
    # Create a new dataframe to hold the processed results
    result_df = []
    
    # Identify duplicate records
    duplicate_indices = []
    for i in range(1, len(df)):
        curr = df.iloc[i]
        prev = df.iloc[i-1]
        
        # Check for duplicate records (same account, product, opportunity ID, and amount)
        if (curr['Account Name'] == prev['Account Name'] and 
            curr['Product Code'] == prev['Product Code'] and 
            curr['Opportunity ID'] == prev['Opportunity ID'] and
            curr['Amount'] == prev['Amount']):
            duplicate_indices.append(i)
    
    # Group by Account Name and Product Code to find single-entry active accounts
    account_groups = df.groupby(['Account Name', 'Product Code'])
    single_entry_active_accounts = []
    
    for (account_name, product_code), group in account_groups:
        if (len(group) == 1 and 
            product_code == '350-0100' and 
            group['Opportunity Type'].iloc[0] == 'Add Products' and
            group['Account Status'].iloc[0] == 'Active'):
            single_entry_active_accounts.append((account_name, product_code))
    
    # First, identify accounts with multiple "Start of Subscription" entries
    # This function finds accounts with multiple Add Products entries for the same product code
    # with no Reduction or Debook entries between them
    accounts_with_multiple_subscriptions = identify_multiple_subscriptions(df)
    
    # Process each row
    i = 0
    while i < len(df):
        current_row = df.iloc[i].copy()
        
        # Handle duplicates
        if i in duplicate_indices:
            current_row['Amount'] = None
            current_row['Note'] = "Duplicate"
            result_df.append(current_row)
            i += 1
            continue
        
        # Special case handling for Copart, Inc - directly in row processing
        if current_row['Account Name'] == 'Copart, Inc' and current_row['Product Code'] == '350-0100':
            # For the second Add Products entry which is $0, mark it as a swap
            if current_row['Amount'] == 0 and current_row['Opportunity Type'] == 'Add Products':
                current_row['Note'] = "Start of Subscription - Swap"
            # For the first entry with $200, keep as Start of Subscription
            elif current_row['Amount'] > 0 and current_row['Opportunity Type'] == 'Add Products':
                current_row['Note'] = "Start of Subscription"
            
        # Process 350-0100 product code (Service Optimization)
        if current_row['Product Code'] == '350-0100':
            # Handle Add Products entries with positive amounts (new subscriptions)
            if current_row['Opportunity Type'] == 'Add Products' and current_row['Amount'] > 0:
                account_name = current_row['Account Name']
                
                # Check if this account has multiple subscriptions and handle accordingly
                if account_name in accounts_with_multiple_subscriptions:
                    subscriptions = accounts_with_multiple_subscriptions[account_name]
                    # Find which subscription number this is based on date
                    for sub_num, sub_info in enumerate(subscriptions, 1):
                        if sub_info['date'] == current_row['Date']:
                            # Set the Note as "Start of Subscription" for all instances
                            current_row['Note'] = "Start of Subscription"
                            
                            # For subsequent subscriptions, modify the Account_Name field
                            if sub_num > 1:
                                # Add suffix to Account_Name for subscriptions after the first
                                current_row['Account Name'] = f"{account_name}_{sub_num}"
                            break
                else:
                    # Normal case - only one subscription
                    current_row['Note'] = "Start of Subscription"
                
                # Check for a subsequent reduction/debook for the same account and product
                reduction_index = -1
                partial_reduction_index = -1
                
                for j in range(i + 1, len(df)):
                    next_row = df.iloc[j]
                    # Skip if this is a marked duplicate
                    if j in duplicate_indices:
                        continue
                        
                    if next_row['Account Name'] == current_row['Account Name'] and next_row['Product Code'] == current_row['Product Code']:
                        # Check for Reduction/Debook opportunity types
                        if next_row['Opportunity Type'] in ['Reduction', 'Debook']:
                            reduction_index = j
                            break
                        # Check for Add Products with negative amount (partial reduction)
                        elif next_row['Opportunity Type'] == 'Add Products' and next_row['Amount'] < 0:
                            partial_reduction_index = j
                            
                    elif next_row['Account Name'] != current_row['Account Name']:
                        # We've moved to a different account, stop looking
                        break
                
                result_df.append(current_row)
                
                # For accounts with multiple subscriptions, we need to handle them specially
                if account_name in accounts_with_multiple_subscriptions and current_row['Account Status'] == 'Active':
                    # We'll handle intermediate entries for these in a separate section
                    # to avoid duplicate processing
                    pass
                
                # Check if this is a single-entry active account that needs to be extended to current month
                # Also include United Mortgage Lending which is a special case
                if current_row['Account Name'] == 'United Mortgage Lending' and current_row['Account Status'] == 'Active':
                    # Always ensure United Mortgage Lending gets the Start of Subscription note
                    current_row['Note'] = 'Start of Subscription'
                
                if (((current_row['Account Name'], current_row['Product Code']) in single_entry_active_accounts and
                     current_row['Account Status'] == 'Active') or 
                    (current_row['Account Name'] == 'United Mortgage Lending' and 
                     current_row['Account Status'] == 'Active')):
                    
                    # Calculate entries from subscription date to current month
                    current_date = current_row['Date']
                    current_month = datetime.now()
                    
                    # Generate intermediate monthly entries
                    next_date = current_date + relativedelta(months=1)
                    while next_date < current_month:
                        # Create new intermediate row
                        intermediate_row = current_row.copy()
                        intermediate_row['Date'] = next_date
                        
                        # Use the same Note suffix for intermediate entries
                        if "Start of Subscription" in current_row['Note']:
                            suffix = current_row['Note'].replace("Start of Subscription", "").strip()
                            if suffix:
                                intermediate_row['Note'] = f"Subscribed Billing {suffix}"
                            else:
                                intermediate_row['Note'] = "Subscribed Billing"
                        else:
                            intermediate_row['Note'] = "Subscribed Billing"
                            
                        result_df.append(intermediate_row)
                        
                        # Move to the next month
                        next_date = next_date + relativedelta(months=1)
                
                # Handle partial reduction case (Example C - Add Products with negative amount)
                if partial_reduction_index != -1:
                    partial_row = df.iloc[partial_reduction_index].copy()
                    
                    # Keep the same numbering for reductions if present
                    if "Start of Subscription" in current_row['Note']:
                        suffix = current_row['Note'].replace("Start of Subscription", "").strip()
                        if suffix:
                            partial_row['Note'] = f"Reduction - Subscribed Billing {suffix}"
                        else:
                            partial_row['Note'] = "Reduction - Subscribed Billing"
                    else:
                        partial_row['Note'] = "Reduction - Subscribed Billing"
                        
                    result_df.append(partial_row)
                    
                    # Skip processing this partial reduction row again
                    i = partial_reduction_index + 1
                    continue
                
                # If we found a reduction/debook entry
                elif reduction_index != -1:
                    reduction_row = df.iloc[reduction_index].copy()
                    current_date = current_row['Date']
                    reduction_date = reduction_row['Date']
                    
                    # Check if there are missing months between add and reduction
                    if current_date != reduction_date:
                        # Generate intermediate monthly entries
                        next_date = current_date + relativedelta(months=1)
                        while next_date < reduction_date:
                            # Create new intermediate row
                            intermediate_row = current_row.copy()
                            intermediate_row['Date'] = next_date
                            
                            # Use the same Note suffix for intermediate entries
                            if "Start of Subscription" in current_row['Note']:
                                suffix = current_row['Note'].replace("Start of Subscription", "").strip()
                                if suffix:
                                    intermediate_row['Note'] = f"Subscribed Billing {suffix}"
                                else:
                                    intermediate_row['Note'] = "Subscribed Billing"
                            else:
                                intermediate_row['Note'] = "Subscribed Billing"
                            
                            # Check if there was a partial reduction before this and adjust amount
                            if partial_reduction_index != -1:
                                partial_row = df.iloc[partial_reduction_index]
                                if partial_row['Date'] < next_date:
                                    # Adjust the amount based on the partial reduction
                                    intermediate_row['Amount'] = float(current_row['Amount']) + float(partial_row['Amount'])
                                    
                            result_df.append(intermediate_row)
                            
                            # Move to the next month
                            next_date = next_date + relativedelta(months=1)
                    
                    # Add the reduction row with appropriate note
                    # Keep the same numbering for reductions if present
                    if "Start of Subscription" in current_row['Note']:
                        suffix = current_row['Note'].replace("Start of Subscription", "").strip()
                        if suffix:
                            reduction_row['Note'] = f"Reduction - End of Subscription {suffix}"
                        else:
                            reduction_row['Note'] = "Reduction - End of Subscription"
                    else:
                        reduction_row['Note'] = "Reduction - End of Subscription"
                        
                    result_df.append(reduction_row)
                    
                    # Skip processing this reduction row again
                    i = reduction_index + 1
                    continue
            
            # Handle "Add Products" with negative amounts (partial reductions) that weren't handled in the previous section
            elif current_row['Opportunity Type'] == 'Add Products' and current_row['Amount'] < 0:
                # Check if there was a previous Add Products entry with positive amount
                add_product_index = -1
                for j in range(i - 1, -1, -1):
                    prev_row = df.iloc[j]
                    if (prev_row['Account Name'] == current_row['Account Name'] and 
                        prev_row['Product Code'] == current_row['Product Code'] and
                        prev_row['Opportunity Type'] == 'Add Products' and
                        prev_row['Amount'] > 0):
                        add_product_index = j
                        break
                    elif prev_row['Account Name'] != current_row['Account Name']:
                        # We've moved to a different account, stop looking
                        break
                
                if add_product_index != -1:
                    # Get the original subscription note to maintain the same numbering
                    original_note = df.iloc[add_product_index]['Note'] if 'Note' in df.iloc[add_product_index] and df.iloc[add_product_index]['Note'] else ""
                    
                    # This is a partial reduction
                    if "Start of Subscription" in original_note:
                        suffix = original_note.replace("Start of Subscription", "").strip()
                        if suffix:
                            current_row['Note'] = f"Reduction - Subscribed Billing {suffix}"
                        else:
                            current_row['Note'] = "Reduction - Subscribed Billing"
                    else:
                        current_row['Note'] = "Reduction - Subscribed Billing"
                else:
                    # Unexpected case, just add a general note
                    current_row['Note'] = "Reduction"
                
                result_df.append(current_row)
                
            # Handle Reduction entries 
            elif current_row['Opportunity Type'] == 'Reduction' and current_row['Amount'] < 0:
                # Check if there was a previous Add Products entry
                add_product_index = -1
                for j in range(i - 1, -1, -1):
                    prev_row = df.iloc[j]
                    if (prev_row['Account Name'] == current_row['Account Name'] and 
                        prev_row['Product Code'] == current_row['Product Code'] and
                        prev_row['Opportunity Type'] == 'Add Products'):
                        add_product_index = j
                        break
                    elif prev_row['Account Name'] != current_row['Account Name']:
                        # We've moved to a different account, stop looking
                        break
                
                if add_product_index != -1:
                    # Get the original subscription note to maintain the same numbering
                    original_note = df.iloc[add_product_index]['Note'] if 'Note' in df.iloc[add_product_index] and df.iloc[add_product_index]['Note'] else ""
                    
                    if abs(df.iloc[add_product_index]['Amount']) > abs(current_row['Amount']):
                        # This is a partial reduction
                        if "Start of Subscription" in original_note:
                            suffix = original_note.replace("Start of Subscription", "").strip()
                            if suffix:
                                current_row['Note'] = f"Reduction - Subscribed Billing {suffix}"
                            else:
                                current_row['Note'] = "Reduction - Subscribed Billing"
                        else:
                            current_row['Note'] = "Reduction - Subscribed Billing"
                    else:
                        # Standard reduction ending subscription
                        if "Start of Subscription" in original_note:
                            suffix = original_note.replace("Start of Subscription", "").strip()
                            if suffix:
                                current_row['Note'] = f"Reduction - End of Subscription {suffix}"
                            else:
                                current_row['Note'] = "Reduction - End of Subscription"
                        else:
                            current_row['Note'] = "Reduction - End of Subscription"
                else:
                    # No matching Add Products entry found
                    current_row['Note'] = "Reduction - End of Subscription"
                
                result_df.append(current_row)
            
            # Handle Debook opportunity type
            elif current_row['Opportunity Type'] == 'Debook':
                # Check if there was a previous Add Products entry to get the numbering
                add_product_index = -1
                for j in range(i - 1, -1, -1):
                    prev_row = df.iloc[j]
                    if (prev_row['Account Name'] == current_row['Account Name'] and 
                        prev_row['Product Code'] == current_row['Product Code'] and
                        prev_row['Opportunity Type'] == 'Add Products'):
                        add_product_index = j
                        break
                    elif prev_row['Account Name'] != current_row['Account Name']:
                        # We've moved to a different account, stop looking
                        break
                
                if add_product_index != -1:
                    # Get the original subscription note to maintain the same numbering
                    original_note = df.iloc[add_product_index]['Note'] if 'Note' in df.iloc[add_product_index] and df.iloc[add_product_index]['Note'] else ""
                    
                    if "Start of Subscription" in original_note:
                        suffix = original_note.replace("Start of Subscription", "").strip()
                        if suffix:
                            current_row['Note'] = f"Reduction - End of Subscription {suffix}"
                        else:
                            current_row['Note'] = "Reduction - End of Subscription"
                    else:
                        current_row['Note'] = "Reduction - End of Subscription"
                else:
                    current_row['Note'] = "Reduction - End of Subscription"
                    
                result_df.append(current_row)
            
            # Add other cases
            else:
                result_df.append(current_row)
        
        # Handle 350-0101 Product Code (On Demand)
        elif current_row['Product Code'] == '350-0101':
            if current_row['Opportunity Type'] == 'Add Products':
                # Set note for On Demand entries
                current_row['Note'] = "On Demand Entry"
                result_df.append(current_row)
            elif current_row['Opportunity Type'] in ['Reduction', 'Debook']:
                current_row['Note'] = "Reduction"
                result_df.append(current_row)
            else:
                # Other cases for 350-0101
                result_df.append(current_row)
        
        # For other product codes, just add the row as is
        else:
            result_df.append(current_row)
            
        i += 1
    
    # Process accounts with multiple subscriptions to add intermediate entries
    for account_name, subscriptions in accounts_with_multiple_subscriptions.items():
        # Sort subscriptions by date
        subscriptions = sorted(subscriptions, key=lambda x: x['date'])
        current_month = datetime.now()
        
        # Iterate through each subscription separately and generate entries up to current month
        for sub_num, sub_info in enumerate(subscriptions, 1):
            current_date = sub_info['date']
            current_amount = sub_info['amount']
            
            # For each subscription, generate entries from its start date to current month
            intermediate_date = current_date + relativedelta(months=1)
            
            # Find the existing row for this subscription to use as template
            # For subscriptions after the first one, we need to look for Account_Name with suffix
            search_account_name = account_name
            if sub_num > 1:
                search_account_name = f"{account_name}_{sub_num}"
                
            base_entries = [r for r in result_df if r['Account Name'] == search_account_name and r['Date'] == current_date]
            
            if base_entries:
                base_row = base_entries[0].copy()
                
                while intermediate_date < current_month:
                    # Create new intermediate row
                    intermediate_row = base_row.copy()
                    intermediate_row['Date'] = intermediate_date
                    intermediate_row['Note'] = 'Subscribed Billing'
                    
                    # Make sure we're using the correct account name with suffix if needed
                    if sub_num > 1:
                        intermediate_row['Account Name'] = f"{account_name}_{sub_num}"
                    
                    result_df.append(intermediate_row)
                    
                    # Move to the next month
                    intermediate_date = intermediate_date + relativedelta(months=1)
    
    # Convert the result list to a dataframe
    result_df = pd.DataFrame(result_df)
    
    # Add special intermediate entries for specific accounts
    result_df = add_special_intermediate_entries(result_df)
    
    # Convert Date back to string format for output
    result_df['Date'] = result_df['Date'].dt.strftime('%m/%d/%Y')
    
    # Replace spaces in column names with underscores for BigQuery compatibility
    result_df.columns = [col.replace(' ', '_') for col in result_df.columns]
    
    # Generate timestamp for the output file
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Create output filename and save the processed data
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_file = f'Quick_Assist_Ledger_Output_{timestamp}.csv'
    output_path = os.path.join(output_dir, output_file)
    
    result_df.to_csv(output_path, index=False)
    print(f"Processing complete. Output saved to {output_path}")
    
    # Calculate and display the number of active subscribers
    # First convert back to datetime for proper filtering
    result_df['Date'] = pd.to_datetime(result_df['Date'])
    
    # Filter for only subscription type accounts (Product_Code = 350-0100)
    subscription_df = result_df[result_df['Product_Code'] == '350-0100']
    
    # Sort by Account_Name and Date in descending order
    subscription_df = subscription_df.sort_values(by=['Account_Name', 'Date'], ascending=[True, False])
    
    # Drop duplicates, keeping the first occurrence (most recent) for each account
    latest_entries = subscription_df.drop_duplicates(subset='Account_Name', keep='first')
    
    # Handle special cases before categorizing
    for idx, row in latest_entries.iterrows():
        account_name = row['Account_Name']
        
        # Case 1: Account status is "Churned" - regardless of current Note value
        if row['Account_Status'] == 'Churned' and row['Opportunity_Type'] == 'Add Products':
            latest_entries.at[idx, 'Note'] = 'Start of Subscription - Churned'
        
        # Case 2: Account status is NaN/null (inactive/disabled accounts) - regardless of current Note value
        elif pd.isna(row['Account_Status']) and row['Opportunity_Type'] == 'Add Products':
            latest_entries.at[idx, 'Note'] = 'Start of Subscription - Inactive'
            
        elif account_name == 'United Mortgage Lending':
            latest_entries.at[idx, 'Note'] = 'Start of Subscription'
            
        # Case 4: Handle Copart, Inc directly and any amount 0 entries with On Demand equivalents
        elif account_name == 'Copart, Inc' and row['Product_Code'] == '350-0100' and row['Amount'] == 0:
            latest_entries.at[idx, 'Note'] = 'Start of Subscription - Swap'
            
        # Case 5: Amount is 0, Account is Active but Note is empty (possible swap to On Demand)
        elif (row['Amount'] == 0 and row['Account_Status'] == 'Active' and 
              row['Opportunity_Type'] == 'Add Products' and (pd.isna(row['Note']) or row['Note'] == '')):
            # Check if there's an On Demand entry with the same Opportunity_ID
            on_demand_entries = result_df[
                (result_df['Account_Name'] == account_name) &
                (result_df['Product_Code'] == '350-0101') &
                (result_df['Opportunity_ID'] == row['Opportunity_ID'])
            ]
            
            if len(on_demand_entries) > 0:
                latest_entries.at[idx, 'Note'] = 'Start of Subscription - Swap'
    
    # Filter for active subscribers based on the specified conditions
    active_subscribers = latest_entries[
        (latest_entries['Opportunity_Type'] == 'Add Products') &
        ((latest_entries['Note'].str.contains('Start of Subscription', na=False)) | 
         (latest_entries['Note'].str.contains('Subscribed Billing', na=False))) &
        (~latest_entries['Note'].str.contains('Churned|Inactive|Swap', na=False)) &
        (latest_entries['Account_Status'] == 'Active')
    ]
    
    # Get the count and the list of active accounts
    num_active_subscribers = active_subscribers['Account_Name'].nunique()
    list_of_active_subscribers = active_subscribers['Account_Name'].unique().tolist()
    
    # Print the results to terminal
    print("\n" + "="*50)
    print(f"Number of active subscribers (Product_Code = 350-0100): {num_active_subscribers}")
    print("\nList of all active subscribers:")
    for account in sorted(list_of_active_subscribers):
        print(f"- {account}")
    print("="*50 + "\n")
    
    # Filter for accounts whose most recent entry shows an ended subscription
    # Now including our special cases as ended subscriptions
    ended_subscription_accounts = latest_entries[
        (latest_entries['Note'].str.contains('Reduction - End of Subscription', na=False)) |
        (latest_entries['Note'].str.contains('Start of Subscription - Churned', na=False)) |
        (latest_entries['Note'].str.contains('Start of Subscription - Inactive', na=False)) |
        (latest_entries['Note'].str.contains('Start of Subscription - Swap', na=False))
    ]

    # Get the count and the list of accounts with ended subscriptions
    num_ended_subscriptions = ended_subscription_accounts['Account_Name'].nunique()
    list_of_ended_subscriptions = ended_subscription_accounts['Account_Name'].unique().tolist()

    # Print the results to terminal
    print("\n" + "="*50)
    print(f"Number of accounts with ended subscriptions (Product_Code = 350-0100): {num_ended_subscriptions}")
    print("\nList of accounts whose most recent entry shows an ended subscription:")
    for account in sorted(list_of_ended_subscriptions):
        print(f"- {account}")
    print("="*50 + "\n")
    
    # Find accounts that are neither active subscribers nor ended subscriptions
    all_accounts = set(latest_entries['Account_Name'].unique())
    active_accounts = set(active_subscribers['Account_Name'].unique())
    ended_accounts = set(ended_subscription_accounts['Account_Name'].unique())
    other_status_accounts = all_accounts - (active_accounts | ended_accounts)

    # Get count and list of these "other status" accounts
    num_other_status = len(other_status_accounts)
    list_other_status = sorted(list(other_status_accounts))

    # Print the results
    print("\n" + "="*50)
    print(f"Number of subscription accounts with other status (Product_Code = 350-0100): {num_other_status}")
    print(f"Total subscription accounts across all categories: {num_active_subscribers + num_ended_subscriptions + num_other_status}")
    print("\nList of subscription accounts with other status and their current state:")
    for account in list_other_status:
        # Get the most recent entry for this account to show its actual status
        account_entry = latest_entries[latest_entries['Account_Name'] == account].iloc[0]
        status_info = f"Note: {account_entry['Note']}, Opportunity_Type: {account_entry['Opportunity_Type']}, Account_Status: {account_entry['Account_Status']}"
        print(f"- {account} ({status_info})")
    print("="*50 + "\n")
    
    # Display information about accounts with multiple subscriptions
    multiple_sub_accounts = [account for account in active_subscribers['Account_Name'].unique() if '_' in account]
    # Get the base account names (without suffixes)
    base_account_names = set()
    for account in multiple_sub_accounts:
        if '_' in account:
            base_name = account.split('_')[0]
            base_account_names.add(base_name)
    
    # Add base account names that exist in the active subscribers list
    for account in active_subscribers['Account_Name'].unique():
        if '_' not in account and any(account + '_' in sub for sub in multiple_sub_accounts):
            base_account_names.add(account)
    
    print("\n" + "="*50)
    print(f"Number of accounts with multiple subscriptions: {len(base_account_names)}")
    print("\nAccounts with multiple subscriptions:")
    for account in sorted(base_account_names):
        # Find all variants of this account (base and with suffixes)
        variants = [acct for acct in active_subscribers['Account_Name'].unique() 
                   if acct == account or acct.startswith(account + '_')]
        print(f"- {account} ({', '.join(sorted(variants))})")
    print("="*50 + "\n")
    
    # Copy the current script to the destination folder
    script_dest_folder = r"C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Revenue Dashboard\QuickAssist_Ledger"
    try:
        # Get the current script path
        current_script_path = os.path.abspath(__file__)
        script_filename = os.path.basename(current_script_path)
        
        # Create the destination folder if it doesn't exist
        if not os.path.exists(script_dest_folder):
            os.makedirs(script_dest_folder)
            print(f"Created directory: {script_dest_folder}")
        
        # Copy the script to the destination
        dest_script_path = os.path.join(script_dest_folder, script_filename)
        shutil.copy2(current_script_path, dest_script_path)
        print(f"Script copied to: {dest_script_path}")
    except Exception as e:
        print(f"Error copying script to destination folder: {e}")
    
    # Move the source file to the archive directory
    try:
        archive_path = os.path.join(archive_dir, os.path.basename(file_path))
        shutil.move(file_path, archive_path)
        print(f"Source file moved to: {archive_path}")
    except Exception as e:
        print(f"Error moving source file to archive: {e}")

def add_special_intermediate_entries(result_df):
    """
    Adds special intermediate date entries for specific accounts as requested.
    This function should be called after all normal processing is complete.
    
    Args:
        result_df (pd.DataFrame): The processed dataframe with all records
        
    Returns:
        pd.DataFrame: The dataframe with special intermediate entries added
    """
    # Make a copy of the dataframe to avoid modifying the original during iteration
    df = result_df.copy()
    
    # Convert Date column back to datetime for processing if it's not already
    if df['Date'].dtype != 'datetime64[ns]':
        df['Date'] = pd.to_datetime(df['Date'])
    
    # New rows to be added
    new_rows = []
    
    # Check if columns have been renamed with underscores
    account_col = 'Account_Name' if 'Account_Name' in df.columns else 'Account Name'
    
    # 1. ADT Solar LLC (fka SUNPRO) - Monthly entries from 2/2/2024 to 8/2/2024
    adt_entries = df[df[account_col] == 'ADT Solar LLC (fka SUNPRO)'].copy()
    if not adt_entries.empty:
        # Find the original entry to use as a template
        template_row = adt_entries.iloc[0].copy()
        
        # Create entries for each month from 2/2/2024 to 8/2/2024
        start_date = datetime(2024, 2, 2)
        end_date = datetime(2024, 8, 2)
        
        current_date = start_date
        while current_date <= end_date:
            new_row = template_row.copy()
            new_row['Date'] = current_date
            new_row['Note'] = "Subscribed Billing"
            new_rows.append(new_row)
            current_date = current_date + relativedelta(months=1)
    
    # 2. Electronic Caregiver - Monthly entries from 4/21/2025 to 6/21/2025 + special entry
    ec_entries = df[df[account_col] == 'Electronic Caregiver'].copy()
    if not ec_entries.empty:
        # Find the original entry to use as a template
        template_row = ec_entries.iloc[0].copy()
        
        # Create entries for each month from 4/21/2025 to 6/21/2025
        start_date = datetime(2025, 4, 21)
        end_date = datetime(2025, 6, 21)
        
        current_date = start_date
        while current_date <= end_date:
            new_row = template_row.copy()
            new_row['Date'] = current_date
            new_row['Note'] = "Subscribed Billing"
            new_rows.append(new_row)
            current_date = current_date + relativedelta(months=1)
        
        # Add special entry for 7/14/2025 with amount 12.58
        special_row = template_row.copy()
        special_row['Date'] = datetime(2025, 7, 14)
        special_row['Amount'] = 12.58
        special_row['Note'] = "Subscribed Billing"
        new_rows.append(special_row)
    
    # 3. Sun Source Energy - Special entry for 12/11/2024 + monthly entries from 1/11/2025 to 6/11/2025
    sse_entries = df[df[account_col] == 'Sun Source Energy'].copy()
    if not sse_entries.empty:
        # Find the original entry to use as a template
        template_row = sse_entries.iloc[0].copy()
        
        # Add special entry for 12/11/2024
        special_row = template_row.copy()
        special_row['Date'] = datetime(2024, 12, 11)
        special_row['Note'] = "Subscribed Billing"
        new_rows.append(special_row)
        
        # Create entries for each month from 1/11/2025 to 6/11/2025 with amount 414.75
        start_date = datetime(2025, 1, 11)
        end_date = datetime(2025, 6, 11)
        
        current_date = start_date
        while current_date <= end_date:
            new_row = template_row.copy()
            new_row['Date'] = current_date
            new_row['Amount'] = 414.75
            new_row['Note'] = "Subscribed Billing"
            new_rows.append(new_row)
            current_date = current_date + relativedelta(months=1)
    
    # Add the new rows to the original dataframe
    if new_rows:
        result_df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
        
        # Sort by Account_Name and Date
        result_df = result_df.sort_values(by=[account_col, 'Date'])
    
    return result_df

def identify_multiple_subscriptions(df):
    """
    Identifies accounts that have:
    1. Multiple "Start of Subscription" entries (Add Products entries for the same product code)
    2. No "Reduction" or "Debook" entries between ANY "Start of Subscription" entries
    3. An "Active" account status
    
    Returns a dictionary with account names as keys and a list of subscription dates as values.
    """
    # Ensure we have the right column names
    date_col = 'Date'
    account_col = 'Account Name'
    product_col = 'Product Code'
    status_col = 'Account Status'
    opportunity_type_col = 'Opportunity Type'
    amount_col = 'Amount'
    
    # Filter for Service Optimization entries (Product Code 350-0100)
    so_df = df[df[product_col] == '350-0100'].copy()
    
    # Sort by Account Name and Date
    so_df = so_df.sort_values(by=[account_col, date_col]).reset_index(drop=True)
    
    # Initialize a dictionary to store results
    accounts_with_multiple_subs = {}
    
    # Get unique account names
    account_names = so_df[account_col].unique()
    
    # Process each account
    for account_name in account_names:
        # Skip special case for Copart, Inc - it's not actually a multiple subscription
        if account_name == "Copart, Inc":
            continue
            
        # Get all entries for this account
        account_entries = so_df[so_df[account_col] == account_name].copy()
        
        # Skip if account is not active
        latest_status = account_entries[status_col].iloc[-1]
        if latest_status != 'Active':
            continue
        
        # Find "Add Products" entries with positive amounts
        start_sub_entries = account_entries[
            (account_entries[opportunity_type_col] == 'Add Products') & 
            (account_entries[amount_col] > 0)
        ]
        
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
                ((entries_between[opportunity_type_col] == 'Add Products') & (entries_between[amount_col] < 0))
            ]
            
            if len(reduction_entries) > 0:
                has_reduction_between = True
                break
        
        # If there are no reductions between any Start of Subscription entries, this is what we're looking for
        if not has_reduction_between:
            # Store the subscription dates and amounts
            subscriptions = []
            for idx, row in start_sub_entries.iterrows():
                subscriptions.append({
                    'date': row[date_col],
                    'amount': row[amount_col]
                })
            
            accounts_with_multiple_subs[account_name] = subscriptions
    
    return accounts_with_multiple_subs

if __name__ == "__main__":
    process_closed_won_opportunities()
