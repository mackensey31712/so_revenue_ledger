import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
import os
from datetime import datetime
import glob
import shutil

def process_closed_won_opportunities():
    # Define directory paths
    input_dir = r"C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Test Files"
    output_dir = r"C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Test Files"
    archive_dir = r"C:\Users\mcgace1\OneDrive - Five9\Documents\Five9\Projects\aii\SO Dashboards\Test Files\Archived"
    
    # Check if directories exist, create if they don't
    for dir_path in [input_dir, output_dir, archive_dir]:
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
            print(f"Created directory: {dir_path}")
    
    # Find specific test file
    file_path = os.path.join(input_dir, "Test_Multiple_Instances.csv")
    
    if not os.path.exists(file_path):
        print(f"Test file not found: {file_path}")
        return
    
    # Process the test file
    process_file(file_path, output_dir, archive_dir)

# Rest of the script remains the same as Quick_Assist_Ledger_V3.py
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
    
    # Identify accounts with multiple "Add Products" entries with no Reduction/Debook in between
    accounts_with_multiple_add_products = {}
    
    for (account_name, product_code), group in account_groups:
        if product_code == '350-0100':
            # Get all "Add Products" entries for this account/product code
            add_products_entries = group[group['Opportunity Type'] == 'Add Products']
            
            # Skip if less than 2 "Add Products" entries
            if len(add_products_entries) < 2:
                continue
                
            # Sort by date
            add_products_entries = add_products_entries.sort_values('Date')
            
            # Check if there are any "Reduction" or "Debook" entries between the "Add Products" entries
            add_product_dates = add_products_entries['Date'].tolist()
            has_reduction_between = False
            
            for i in range(1, len(add_product_dates)):
                prev_date = add_product_dates[i-1]
                curr_date = add_product_dates[i]
                
                # Check for "Reduction" or "Debook" entries between these dates
                reduction_entries = group[
                    (group['Date'] > prev_date) & 
                    (group['Date'] < curr_date) & 
                    (group['Opportunity Type'].isin(['Reduction', 'Debook']))
                ]
                
                if len(reduction_entries) > 0:
                    has_reduction_between = True
                    break
            
            # If no reductions between add products and account is active, add to our list
            if not has_reduction_between and group['Account Status'].iloc[-1] == 'Active':
                accounts_with_multiple_add_products[account_name] = add_products_entries.index.tolist()
    
    # Print accounts with multiple "Add Products" entries
    if accounts_with_multiple_add_products:
        print("\n" + "="*50)
        print(f"Accounts with multiple Add Products entries and no Reduction/Debook in between: {len(accounts_with_multiple_add_products)}")
        for account, indices in accounts_with_multiple_add_products.items():
            print(f"- {account}: {len(indices)} entries")
        print("="*50 + "\n")
    
    # Create a dictionary to track account name modifications
    account_name_counters = {}
    
    # Process each row
    i = 0
    while i < len(df):
        current_row = df.iloc[i].copy()
        current_account = current_row['Account Name']
        
        # Handle duplicates
        if i in duplicate_indices:
            current_row['Amount'] = None
            current_row['Note'] = "Duplicate"
            result_df.append(current_row)
            i += 1
            continue
        
        # Check if this is a multi-entry account that needs to be renamed
        if (current_account in accounts_with_multiple_add_products and 
            current_row['Product Code'] == '350-0100' and 
            current_row['Opportunity Type'] == 'Add Products'):
            
            # Initialize counter for this account if not already done
            if current_account not in account_name_counters:
                account_name_counters[current_account] = 0
            
            # Increment counter for this account
            account_name_counters[current_account] += 1
            
            # Append counter to account name if this is not the first entry
            if account_name_counters[current_account] > 1:
                current_row['Account Name'] = f"{current_account}_{account_name_counters[current_account]}"
                
        # Process 350-0100 product code (Service Optimization)
        if current_row['Product Code'] == '350-0100':
            # Handle Add Products entries with positive amounts (new subscriptions)
            if current_row['Opportunity Type'] == 'Add Products' and current_row['Amount'] > 0:
                # Check for a subsequent reduction/debook for the same account and product
                reduction_index = -1
                partial_reduction_index = -1
                
                # Set the initial note
                current_row['Note'] = "Start of Subscription"
                result_df.append(current_row)
                
                # For multiple entry accounts, always generate entries from subscription date to current month
                original_account_name = current_row['Account Name']
                if "_" in original_account_name and any(original_account_name.startswith(acct) for acct in accounts_with_multiple_add_products):
                    base_name = original_account_name.split("_")[0]
                    if base_name in accounts_with_multiple_add_products:
                        # Calculate entries from subscription date to current month
                        current_date = current_row['Date']
                        current_month = datetime.now()
                        
                        # Generate intermediate monthly entries
                        next_date = current_date + relativedelta(months=1)
                        while next_date < current_month:
                            # Create new intermediate row
                            intermediate_row = current_row.copy()
                            intermediate_row['Date'] = next_date
                            intermediate_row['Note'] = "Subscribed Billing"
                            result_df.append(intermediate_row)
                            
                            # Move to the next month
                            next_date = next_date + relativedelta(months=1)
                        
                        # Move to the next row
                        i += 1
                        continue
                
                # For regular single-entry cases, continue with the normal logic
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
                
                # Check if this is a single-entry active account that needs to be extended to current month
                # Also include United Mortgage Lending which is a free subscription (Amount = 0)
                if (((current_row['Account Name'], current_row['Product Code']) in single_entry_active_accounts and
                     current_row['Account Status'] == 'Active') or 
                    (current_row['Account Name'] == 'United Mortgage Lending' and 
                     current_row['Amount'] == 0 and 
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
                        intermediate_row['Note'] = "Subscribed Billing"
                        result_df.append(intermediate_row)
                        
                        # Move to the next month
                        next_date = next_date + relativedelta(months=1)
                
                # Handle partial reduction case (Example C - Add Products with negative amount)
                if partial_reduction_index != -1:
                    partial_row = df.iloc[partial_reduction_index].copy()
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
                    # This is a partial reduction
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
                
                if add_product_index != -1 and abs(df.iloc[add_product_index]['Amount']) > abs(current_row['Amount']):
                    # This is a partial reduction
                    current_row['Note'] = "Reduction - Subscribed Billing"
                else:
                    # Standard reduction ending subscription
                    current_row['Note'] = "Reduction - End of Subscription"
                
                result_df.append(current_row)
            
            # Handle Debook opportunity type
            elif current_row['Opportunity Type'] == 'Debook':
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
    
    # Convert the result list to a dataframe
    result_df = pd.DataFrame(result_df)
    
    # Convert Date back to string format for output
    result_df['Date'] = result_df['Date'].dt.strftime('%m/%d/%Y')
    
    # Replace spaces in column names with underscores for BigQuery compatibility
    result_df.columns = [col.replace(' ', '_') for col in result_df.columns]
    
    # Generate timestamp for the output file
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Create output filename and save the processed data
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_file = f'Test_Result_{timestamp}.csv'
    output_path = os.path.join(output_dir, output_file)
    
    result_df.to_csv(output_path, index=False)
    print(f"Processing complete. Output saved to {output_path}")
    
    # Print complete CSV content for verification
    print("\nFull output CSV content:")
    print(result_df.to_string())
    
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
        
        # Case 3: Handle specific accounts by name
        elif account_name == 'Copart, Inc':
            latest_entries.at[idx, 'Note'] = 'Start of Subscription - Swap'
            
        elif account_name == 'United Mortgage Lending':
            latest_entries.at[idx, 'Note'] = 'Start of Subscription'
            
        # Case 4: Amount is 0, Account is Active but Note is empty (possible swap to On Demand)
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
    
    # Get base account names (without _2, _3 suffixes) for accounts with multiple entries
    base_account_names = {}
    for account in latest_entries['Account_Name'].tolist():
        if '_' in account:
            parts = account.split('_')
            if parts[-1].isdigit() and len(parts) > 1:
                base_name = '_'.join(parts[:-1])
                if base_name not in base_account_names:
                    base_account_names[base_name] = []
                base_account_names[base_name].append(account)
    
    # Filter for active subscribers based on the specified conditions
    active_subscribers = latest_entries[
        (latest_entries['Opportunity_Type'] == 'Add Products') &
        (latest_entries['Note'].isin(['Start of Subscription', 'Subscribed Billing'])) &
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
        (latest_entries['Note'] == 'Reduction - End of Subscription') |
        (latest_entries['Note'] == 'Start of Subscription - Churned') |
        (latest_entries['Note'] == 'Start of Subscription - Inactive') |
        (latest_entries['Note'] == 'Start of Subscription - Swap')
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
    
    # Print information about accounts with multiple entries
    if base_account_names:
        print("\n" + "="*50)
        print(f"Accounts with multiple entries that were renamed:")
        for base_name, variants in base_account_names.items():
            print(f"- {base_name}: {len(variants) + 1} instances (Original plus {', '.join(variants)})")
        print("="*50 + "\n")

if __name__ == "__main__":
    process_closed_won_opportunities()
