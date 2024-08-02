import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the excel file into a DataFrame
df_repayment = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Repayment')

# Display the first 5 rows
print(df_repayment.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df_repayment.info())

# Convert `Date` column to datetime
df_repayment['Date'] = pd.to_datetime(df_repayment['Date'], format='%Y-%m-%d')

# Extract month and year from `Date`
df_repayment['Month'] = df_repayment['Date'].dt.month_name()
df_repayment['Year'] = df_repayment['Date'].dt.year

# Aggregate the data by `Customer`, `Month` and `Year` and sum the `Repayment`
df_agg = df_repayment.groupby(['Customer', 'Month', 'Year'])['Repayment'].sum().reset_index()

# Sort the dataframe by `Month`, `Year` and `Repayment` in descending order
df_agg = df_agg.sort_values(by=['Year', 'Month', 'Repayment'], ascending=[True, True, False])

# Group by `Month` and `Year` and select the top 10 rows based on `Repayment`
df_top10 = df_agg.groupby(['Month', 'Year']).head(10)

# Write the final dataframe to an excel file
df_top10.to_excel('top_10_customers_monthly_repayment.xlsx', index=False)
