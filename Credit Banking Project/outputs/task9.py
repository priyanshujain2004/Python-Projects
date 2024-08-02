import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the excel files into Pandas Dataframes
df_spend = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Spend')
df_repayment = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Repayment')

# Display the first 5 rows
print("First 5 rows of Spend Data:")
print(df_spend.head().to_markdown(index=False,numalign="left", stralign="left"))
print("\nFirst 5 rows of Repayment Data:")
print(df_repayment.head().to_markdown(index=False,numalign="left", stralign="left"))

# Print the column names and their data types
print("\nSpend Data column names and types:")
print(df_spend.info())
print("\nRepayment Data column names and types:")
print(df_repayment.info())

# Convert `Date` to datetime
df_spend['Date'] = pd.to_datetime(df_spend['Date'])
df_repayment['Date'] = pd.to_datetime(df_repayment['Date'])

# Extract year and month from `Date`
for df in [df_spend, df_repayment]:
  df['Year'] = df['Date'].dt.year
  df['Month'] = df['Date'].dt.month_name()

# Aggregate spend by year and month
df_spend_agg = df_spend.groupby(['Year', 'Month'])['Spend'].sum().reset_index().rename(columns={'Spend':'Total Monthly Spend'})

# Aggregate repayment by year and month
df_repayment_agg = df_repayment.groupby(['Year', 'Month'])['Repayment'].sum().reset_index().rename(columns={'Repayment':'Total Monthly Repayment'})

# Merge the two aggregated dataframes
df_merged = df_spend_agg.merge(df_repayment_agg, on=['Year', 'Month'], how='inner')

# Calculate profit
df_merged['Profit'] = (df_merged['Total Monthly Spend'] - df_merged['Total Monthly Repayment']) * 0.029

# Create pivot table
pivot_table = df_merged.pivot(index='Year', columns='Month', values='Profit').fillna(0)

# Write pivot table to excel
pivot_table.to_excel("bank's monthly profit.xlsx")

