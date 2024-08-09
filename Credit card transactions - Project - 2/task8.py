import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV file into a DataFrame
df_credit_card = pd.read_csv('Credit card transactions - Project - 2.csv')

# Display the first 5 rows
print(df_credit_card.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df_credit_card.info())

# Convert `Date` column to datetime
df_credit_card['Date'] = pd.to_datetime(df_credit_card['Date'], format='%d-%b-%y')

# Extract day of week from `Date`
df_credit_card['Day_of_Week'] = df_credit_card['Date'].dt.dayofweek

# Filter data for weekends (Saturday and Sunday)
df_weekend = df_credit_card[df_credit_card['Day_of_Week'].isin([5, 6])].copy()

# Group filtered data on `City` and find the sum of `Amount` and count of transactions
df_result = df_weekend.groupby('City').agg(Total_Spend=('Amount', 'sum'),
                                           Total_Transactions=('Amount', 'count')
                                           ).reset_index()

# Calculate ratio of `Total_Spend` to `Total_Transactions` for each `City`
df_result['Ratio'] = df_result['Total_Spend'] / df_result['Total_Transactions']

# Sort in descending order of ratio and pick top row
df_result.sort_values('Ratio', ascending=False, inplace=True)
df_result = df_result.head(1)

# Write the output to excel file 'spend_transaction_ratio_weekends.xlsx'
df_result.to_excel('(task8)spend_transaction_ratio_weekends.xlsx', index=False)

# Print success message
print("Data exported to '(task8)spend_transaction_ratio_weekends.xlsx' successfully!")
