import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV file into a DataFrame
df_transaction = pd.read_csv('Credit card transactions - Project - 2.csv')

# Display the first 5 rows
print(df_transaction.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df_transaction.info())

# Filter the data on `Card Type` = 'Gold'
df_gold = df_transaction[df_transaction['Card Type'] == 'Gold'].copy()

# Group the filtered data on `City` and calculate the sum of `Amount`
df_city_spend = df_gold.groupby('City')['Amount'].sum().reset_index()

# Rename the column to `City Spend`
df_city_spend = df_city_spend.rename(columns={'Amount': 'City Spend'})

# Calculate the total spend for 'Gold' card type
total_spend_gold = df_gold['Amount'].sum()

# Calculate the percentage spend for each city
df_city_spend['Percentage Spend'] = (df_city_spend['City Spend'] / total_spend_gold) * 100

# Sort the dataframe in ascending order of `Percentage Spend` and pick the top row
lowest_percentage_city = df_city_spend.sort_values('Percentage Spend').head(1)

# Write the output dataframe to excel sheet
lowest_percentage_city.to_excel('(task4)lowest_percentage_spend_city.xlsx', index=False)

print("DataFrame is written to Excel File successfully.")
