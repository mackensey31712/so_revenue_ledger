# Multi-Entry Accounts Analysis

## Summary

Based on our analysis of the Service Optimization subscription data, we identified **5 accounts** that have:
1. Multiple "Add Products" entries for the same product code (350-0100)
2. No "Reduction" or "Debook" entries after their latest "Add Products" entry
3. An "Active" account status

## Identified Accounts

1. **Canada Custom Autoworks**
   - First Add Date: 04/17/2025
   - Latest Add Date: 08/05/2025
   - Number of entries: 2
   - Time between entries: ~3.5 months

2. **Copart, Inc**
   - First Add Date: 09/09/2024
   - Latest Add Date: 09/26/2024
   - Number of entries: 2
   - Time between entries: ~17 days
   - *Note: This account is already handled as a special case in Quick_Assist_Ledger_V2.py*

3. **Science Based Health**
   - First Add Date: 11/26/2024
   - Latest Add Date: 12/04/2024
   - Number of entries: 2
   - Time between entries: ~8 days

4. **Sky Auto Finance**
   - First Add Date: 09/23/2024
   - Latest Add Date: 06/25/2025
   - Number of entries: 2
   - Time between entries: ~9 months

5. **Solis Health Plans**
   - First Add Date: 12/05/2023
   - Latest Add Date: 02/05/2025
   - Number of entries: 4
   - Multiple entries over ~14 months

## Analysis

The existence of multiple "Add Products" entries without corresponding "Reduction" entries between them suggests several possible scenarios:

1. **Subscription Upgrades**: Accounts may be upgrading their subscriptions without formal reductions of previous subscriptions.

2. **Administrative Adjustments**: Some entries might represent administrative changes rather than actual new subscriptions.

3. **Billing Adjustments**: Multiple entries could indicate changes in billing arrangements.

4. **Data Entry Inconsistencies**: Some entries might be due to inconsistent data entry practices.

## Recommendations

### Short-term

1. **Update Quick_Assist_Ledger_V2.py**: Modify the script to handle multi-entry accounts similar to how it already handles Copart, Inc. The script should:
   - Identify all accounts with multiple Add Products entries
   - Generate intermediate monthly entries from the latest Add Products date to the current month
   - Use the same Amount value as the latest Add Products entry

2. **Data Validation**: Review these specific accounts with the Sales team to understand why they have multiple Add Products entries and confirm the correct subscription amounts.

### Long-term

1. **Standardize Process**: Work with Sales Operations to standardize how subscription changes are recorded in Salesforce:
   - Use "Reduction" entries consistently before adding new subscriptions
   - Develop clear guidelines for handling subscription upgrades vs. new subscriptions

2. **Enhanced Reporting**: Develop a specialized report for accounts with non-standard subscription patterns to ensure they're monitored regularly.

## Implementation Plan

1. Create a new version of Quick_Assist_Ledger.py that identifies and correctly processes multi-entry accounts
2. Test the updated script with the identified accounts
3. Document the additional logic in the code and README
4. Deploy to production with monitoring for these specific accounts

This approach will ensure accurate subscription tracking for all account types while providing a path toward more standardized data entry practices in the future.
