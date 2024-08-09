import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV file into a DataFrame
df_credit_card = pd.read_csv('Credit card transactions - Project - 2.csv')

# Display the first 5 rows
print(df_credit_card.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df_credit_card.info())

# Group data on `City` and find the count of transactions
df_city_transaction_counts = df_credit_card.groupby('City')['Amount'].count()

# Filter out cities with less than 500 transactions
df_city_transaction_counts = df_city_transaction_counts[df_city_transaction_counts >= 500]

# Filter the original dataframe to include only cities with at least 500 transactions
df_credit_card_filtered = df_credit_card[df_credit_card['City'].isin(df_city_transaction_counts.index)]

# Convert `Date` column to datetime
df_credit_card_filtered['Date'] = pd.to_datetime(df_credit_card_filtered['Date'], format='%d-%b-%y')

# Sort in ascending order of `Date`
df_credit_card_filtered.sort_values('Date', inplace=True)

# Group data on `City` and find the first and 500th transaction date
df_result = df_credit_card_filtered.groupby('City').agg(
    First_Transaction_Date=('Date', 'first'),
    Transaction_500_Date=('Date', lambda x: x.iloc[499])
).reset_index()

# Calculate difference in days
df_result['Days_Taken'] = (
    df_result['Transaction_500_Date'] - df_result['First_Transaction_Date']
).dt.days

# Sort in ascending order of days taken and pick top row
df_result.sort_values('Days_Taken', inplace=True)
df_result = df_result.head(1)

# Write the output to excel
df_result.to_excel('(task9)least_days_to_500_transactions.xlsx', index=False)

# Print success message
print("Data exported to '(task9)least_days_to_500_transactions.xlsx' successfully!")
