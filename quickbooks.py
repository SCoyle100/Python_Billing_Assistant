import pandas as pd

from quickbooks import QuickBooks
from quickbooks.objects.account import Account
from quickbooks.objects.journalentry import JournalEntry, JournalEntryLineDetail, Line

# Load data from the Excel file
excel_file_path = 'path/to/RSC_Nov_2023_Financials_Blank_Months.xlsx'
df = pd.read_excel(excel_file_path)

# QuickBooks OAuth2 configuration
client_id = 'YOUR_CLIENT_ID'
client_secret = 'YOUR_CLIENT_SECRET'
access_token = 'YOUR_ACCESS_TOKEN'
refresh_token = 'YOUR_REFRESH_TOKEN'
realm_id = 'YOUR_REALM_ID'

# Initialize QuickBooks client
qb_client = QuickBooks(
    sandbox=True,  # Set to False for production
    client_id=client_id,
    client_secret=client_secret,
    access_token=access_token,
    refresh_token=refresh_token,
    company_id=realm_id
)

# Function to create or update accounts in QuickBooks
def create_or_update_account(name, account_type):
    # Check if the account already exists
    existing_account = Account.filter(Name=name, qb=qb_client)
    if existing_account:
        return existing_account[0]
    
    # Create a new account if it doesn't exist
    account = Account()
    account.Name = name
    account.AccountType = account_type
    account.save(qb=qb_client)
    return account

# Map account names to QuickBooks account types
account_type_mapping = {
    'Income': 'Income',
    'Direct Costs': 'Cost of Goods Sold',
    'Operating Expenses': 'Expense',
    'Net Income (Loss)': 'Equity'
}

# Loop through each row in the DataFrame and add entries to QuickBooks
for _, row in df.iterrows():
    account_name = row['Account']
    account_type = account_type_mapping.get(account_name.split()[0], 'Expense')  # Default to Expense if not mapped
    
    # Create or retrieve the account in QuickBooks
    account = create_or_update_account(account_name, account_type)
    
    # Iterate over each month and add a journal entry line if there's a value
    for month in ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']:
        amount = row[month]
        
        if pd.notnull(amount):  # Only proceed if there's an actual amount
            # Create a journal entry line
            line = Line()
            line.Amount = abs(amount)
            line.DetailType = "JournalEntryLineDetail"
            line.JournalEntryLineDetail = JournalEntryLineDetail()
            line.JournalEntryLineDetail.AccountRef = account.to_ref()
            
            # Credit or Debit based on amount sign
            if amount < 0:
                line.JournalEntryLineDetail.PostingType = "Credit"
            else:
                line.JournalEntryLineDetail.PostingType = "Debit"
            
            # Create and save the journal entry
            journal_entry = JournalEntry()
            journal_entry.Line.append(line)
            journal_entry.save(qb=qb_client)

print("Data successfully populated into QuickBooks.")
