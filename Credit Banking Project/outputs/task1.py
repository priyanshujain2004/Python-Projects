import pandas as pd

# Read the Excel file into a DataFrame, specifying the sheet name
df_spend = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Spend')

# Display the first 5 rows
print(df_spend.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df_spend.info())

# Convert `Date` column to datetime
df_spend['Date'] = pd.to_datetime(df_spend['Date'])

# Extract year and month from `Date`
df_spend['Year'] = df_spend['Date'].dt.year
df_spend['Month'] = df_spend['Date'].dt.month

# Group by `Customer`, `Year` and `Month` columns, sum the `Spend` column
df_agg = df_spend.groupby(['Customer', 'Year', 'Month'])['Spend'].sum().reset_index()

# Combine the `Year` and `Month` columns into a single column `Year_Month` in the format 'YYYY-MM'
df_agg['Year_Month'] = pd.to_datetime(df_agg['Year'].astype(str) + '-' + df_agg['Month'].astype(str)).dt.to_period('M')

# Pivot the table to have `Year_Month` as index and `Customer` as columns and values as `Spend`. Fill missing values with 0.
df_pivot = df_agg.pivot(index='Year_Month', columns='Customer', values='Spend').fillna(0)

# Write the final dataframe to an excel file called 'customer_monthly_spend.xlsx'
df_pivot.to_excel('customer_monthly_spend.xlsx')
