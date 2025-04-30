import pandas as pd
import math
import numpy as np
import openpyxl as op
from openpyxl.worksheet.datavalidation import DataValidation
import json
import requests
import datetime
from dateutil.relativedelta import relativedelta
from datetime import timedelta

# INSTRUCTIONS
# Download AMEX Excel
# Process it by placing name in code below
# It will stack to all previous transactions and create a new file for it

# Read new AMEX transaction file and transform

amex_file = 'AMEX CC/Mar 2025.xls'    

# Read excel sheet of last month transaction which was exported from AMEX website
raw_df = pd.read_excel(amex_file)
raw_df

# Drop unneeded columns
raw_df = raw_df.drop(['Transaction Details: ', 'Unnamed: 4','Unnamed: 5','Unnamed: 6','Unnamed: 8','Unnamed: 9'], axis=1)
raw_df

# Drop unneeded rows
raw_df = raw_df.drop([0,1,2,3,4,5,6,7,8,9])
raw_df

# Reset index
raw_df = raw_df.reset_index(drop=True)
raw_df

# Replace column names and reset index
raw_df.columns = raw_df.iloc[0]
raw_df = raw_df.drop([0])
raw_df = raw_df.reset_index(drop=True)
raw_df

# Add category and subcategory columns
raw_df['Category'] = np.nan
raw_df['Subcategory'] = np.nan
raw_df

# Rearrange columns and Rename columns to make like training data
rearranged_cols = ['Date','Amount','Description','Merchant','Category','Subcategory']
raw_df = raw_df[rearranged_cols]
raw_df = raw_df.rename(columns={"Description": "Merchant Name", "Merchant": "Note"})
raw_df

raw_df['Account'] = 'AMEX'

# Read new BMO Debit Card transactions and transform

bmo_db_file = 'BMO DC/Mar 2025.csv'

from io import StringIO
file_path = bmo_db_file

with open(file_path, 'r') as file:
    # Read the content of the file
    content = file.read()

# Read BMO csv
#bmo_db_df = pd.read_csv(bmo_db_file)
#bmo_db_df

from io import StringIO
file_path = bmo_db_file

with open(file_path, 'r') as file:
    # Read the content of the file
    content = file.read()

# Remove top row
content = content[86:]
content
bmo_db_df = pd.read_csv(StringIO(content))
bmo_db_df

# Convert Date
bmo_db_df['Date Posted'] = pd.to_datetime(bmo_db_df['Date Posted'], format='%Y%m%d')
bmo_db_df

# Drop unneeded columns
bmo_db_df = bmo_db_df.drop(['First Bank Card','Transaction Type'],axis=1)
bmo_db_df

# Rename Columns
bmo_db_df = bmo_db_df.rename(columns={"Date Posted": "Date", " Transaction Amount": "Amount", "Description":"Merchant Name"})
bmo_db_df

# Add category and subcategory columns
bmo_db_df['Category'] = np.nan
bmo_db_df['Subcategory'] = np.nan
bmo_db_df['Account'] = 'BMO Debit'
bmo_db_df

# Read new BMO Business transactions and transform

bmo_business_file = 'BMO Business/Mar 2025.csv'

from io import StringIO
file_path = bmo_business_file

with open(file_path, 'r') as file:
    # Read the content of the file
    content = file.read()

# Remove top row
content = content[98:]
content
bmo_business_df = pd.read_csv(StringIO(content))
bmo_business_df

# Convert Date
bmo_business_df['Date Posted'] = pd.to_datetime(bmo_business_df['Date Posted'], format='%Y%m%d')
bmo_business_df

# Drop unneeded columns
bmo_business_df = bmo_business_df.drop(['First Bank Card','Transaction Type'],axis=1)
bmo_business_df

# Rename Columns
bmo_business_df = bmo_business_df.rename(columns={"Date Posted": "Date", " Transaction Amount": "Amount", "Description":"Merchant Name"})
bmo_business_df

# Add category and subcategory columns
bmo_business_df['Category'] = np.nan
bmo_business_df['Subcategory'] = np.nan
bmo_business_df['Account'] = 'BMO Business'
bmo_business_df

# Read new BMO Credit Card transactions and transform

bmo_cc_file = 'BMO CC/Mar 2025.csv'

# # Read BMO csv
from io import StringIO
file_path = bmo_cc_file

with open(file_path, 'r') as file:
    # Read the content of the file
    content = file.read()
    
# Remove top row
content = content[50:]
content
bmo_cc_df = pd.read_csv(StringIO(content))
bmo_cc_df

# Convert Date
bmo_cc_df['Posting Date'] = pd.to_datetime(bmo_cc_df['Posting Date'], format='%Y%m%d')
bmo_cc_df

# Drop unneeded columns
bmo_cc_df = bmo_cc_df.drop(['Item #','Card #','Transaction Date'],axis=1)
bmo_cc_df

# Rename Columns
bmo_cc_df = bmo_cc_df.rename(columns={"Posting Date": "Date", "Transaction Amount": "Amount", "Description":"Merchant Name"})
bmo_cc_df

# Add category and subcategory columns
bmo_cc_df['Category'] = np.nan
bmo_cc_df['Subcategory'] = np.nan
bmo_cc_df['Account'] = 'BMO Credit'
bmo_cc_df

# Read Historical transactions and stack with new transactions

# Read historical data so that we can stack new data with it
training_df = pd.read_excel('Archive/Feb 2025.xlsx')
training_df

# Stack new data with training data
df = pd.concat([training_df, bmo_db_df])
df = pd.concat([df, bmo_cc_df])
df = pd.concat([df, raw_df])
df = pd.concat([df, bmo_business_df])
df

# Normalize dates so that we can sort by date
df['Date'] = pd.to_datetime(df['Date'])
df

# Sort by date
df = df.sort_values(by='Date')
df

# Reset index
df = df.reset_index(drop=True)
df

# Assign categories and subcategories to new transactions
transaction_list = []
for i in range(len(df)):
    transaction = df.iloc[i].to_dict()
    if type(transaction['Category']) == str:
        transaction_list.append(transaction)

# Iterate through df and convert transactions to dictionaries
for i in range(len(df)):
    transaction = df.iloc[i].to_dict()
    # If the Category in the transaction is empty(float), look through the transaction list
    if type(transaction['Category']) == float:
        for j in range(len(transaction_list)):
            # If the transaction merchant name matches, assign the category and subcategory
            if transaction['Merchant Name'] == transaction_list[j]['Merchant Name']:
                df.iloc[i,4] = transaction_list[j]['Category']
                df.iloc[i,5] = transaction_list[j]['Subcategory']
                break

df

# Change column B and C's values to integers
df = df.astype({'Note': str})

# Categorize as Food if Note cointains RESTAURANT and Category is Empty
def update_category(row):
    if pd.isna(row['Category']) and 'RESTAURANT' in row['Note']:
        return 'Food'
    return row['Category']

df['Category'] = df.apply(update_category, axis=1)

df

# Remove Notes larger than 30 characters
def update_note(row):
    if pd.isna(row['Note']) == False and len(row['Note']) > 30:
        return np.nan
    return row['Note']

df['Note'] = df.apply(update_note, axis=1)

df

df = df.reset_index(drop=True)
df

# Remove dollar signs and commas for strings
df['Amount'] = df['Amount'].apply(lambda x: str(x).replace('$', '').replace(',', '') if isinstance(x, str) else x)

# Convert the entire column to numeric
df['Amount'] = pd.to_numeric(df['Amount'])

df

# Transformation code which contains private information XXXX for privacy

# Remove rows where 'Merchant Name' contains 'XXXX'
df = df.loc[~df['XXXX'].str.contains('XXXX', na=False)]

# Remove rows where 'Merchant Name' contains 'XXXX'
df = df.loc[~df['XXXX'].str.contains('XXXX', na=False)]

# Remove rows where 'Merchant Name' contains 'XXXX'
df = df.loc[~df['XXXX'].str.contains('XXXX', na=False)]

# Remove rows where 'Merchant Name' contains 'XXXX'
df = df.loc[~df['XXXX'].str.contains('XXXX', na=False)]

# Remove rows where 'Merchant Name' contains 'XXXX'
df = df.loc[~df['XXXX'].str.contains('XXXX', na=False)]

# Remove rows where 'Merchant Name' contains 'XXXX' and 'Account' is 'XXXX'
df = df.loc[~((df['XXXX'].str.contains('XXXX', na=False)) & (df['Account'] == 'AMEX'))]

# Change negative values to positive in the 'Amount' column where 'Account' is 'Merchant Name' and 'Amount' is negative
df.loc[(df['Account'] == 'BMO Debit') & (df['Amount'] < 0), 'Amount'] = df['Amount'].abs()

# Change negative values to positive in the 'Amount' column where 'Account' is 'Merchant Name' and 'Amount' is negative
df.loc[(df['Account'] == 'BMO Business') & (df['Amount'] < 0), 'Amount'] = df['Amount'].abs()

# Set 'Category' and 'Subcategory' to 'Amazon' if 'Merchant Bame' contains 'AMAZON.CA' or 'AMZN MKTP'
df.loc[df['XXXX'].str.contains('AMAZON.CA|AMZN MKTP', na=False), ['Category', 'Subcategory']] = 'Amazon'

# Calculate the start and end of last month
today = datetime.datetime.today()
first_day_this_month = today.replace(day=1,hour=0,minute=0,second=0,microsecond=0)
last_day_last_month = first_day_this_month - timedelta(days=1)
first_day_last_month = last_day_last_month.replace(day=1)

# Create the 'last_month' column based on the date range
df['last_month'] = df['Date'].apply(lambda x: 1 if first_day_last_month <= x <= last_day_last_month else 0)

# Transform date to string format because it's easier to use in Power BI
df['Date'] = df['Date'].astype(str)

df.to_csv('Mar 2025.csv', index=False)