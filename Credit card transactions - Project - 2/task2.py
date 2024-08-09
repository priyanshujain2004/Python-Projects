import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV file into a DataFrame
df_transaction = pd.read_csv('Credit card transactions - Project - 2.csv')

# Display the first 5 rows
print(df_transaction.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df_transaction.info())

# Convert `Date` column to datetime
df_transaction['Date'] = pd.to_datetime(df_transaction['Date'], format='%d-%b-%y')

# Extract year and month from `Date`
df_transaction['Year'] = df_transaction['Date'].dt.year
df_transaction['Month'] = df_transaction['Date'].dt.month_name()

# group by `Year` and `Month` and sum up `Amount` field
df_grouped = df_transaction.groupby(['Year', 'Month'])['Amount'].sum().reset_index()

# sort grouped dataframe in descending order of `Amount`
df_grouped = df_grouped.sort_values(by='Amount',ascending=False)

# drop duplicates from grouped dataframe on `Year` keeping the first value
df_result = df_grouped.drop_duplicates(subset=['Year'], keep='first').reset_index(drop=True)

# merge grouped dataframe with original dataframe on `Year` and `Month`
df_merged = pd.merge(df_result, df_transaction, on=['Year', 'Month'])

# group merged dataframe on `Year`, `Month` and `Card Type` and sum up `Amount_y`
df_final = df_merged.groupby(['Year', 'Month', 'Card Type'])['Amount_y'].sum().reset_index()

# print the output
print(df_final.round(2).to_markdown(index=False, numalign="left", stralign="left"))

# write to excel
df_final.to_excel("(task2)month_year_card_type_spend.xlsx", index=False)
