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

import calendar

# Convert `Date` to datetime in both dataframes
df_spend['Date'] = pd.to_datetime(df_spend['Date'])
df_repayment['Date'] = pd.to_datetime(df_repayment['Date'])

# Extract year and month from `Date` in both dataframes
for df in [df_spend, df_repayment]:
  df['Year'] = df['Date'].dt.year
  df['Month'] = df['Date'].dt.month

# Aggregate spend data
df_spend_agg = df_spend.groupby(['Customer', 'Year', 'Month'])['Spend'].sum().reset_index().rename(columns={'Spend': 'Total Spend'})

# Aggregate repayment data
df_repayment_agg = df_repayment.groupby(['Customer', 'Year', 'Month'])['Repayment'].sum().reset_index().rename(columns={'Repayment': 'Total Repayment'})

# Merge the aggregated dataframes
df_merged = pd.merge(df_spend_agg, df_repayment_agg, on=['Customer', 'Year', 'Month'], how='outer').fillna(0)

# Calculate Due Amount, Profit, and Net Payable Amount
df_merged['Due Amount'] = df_merged['Total Spend'] - df_merged['Total Repayment']
df_merged['Profit'] = 0.029 * df_merged['Due Amount']
df_merged['Net Payable Amount'] = df_merged['Due Amount'] + df_merged['Profit']

# Convert month number to month name
df_merged['Month'] = df_merged['Month'].apply(lambda x: calendar.month_name[x])

# Create a pivot table
pivot_table = df_merged.pivot_table(index=['Customer', 'Year', 'Month'], values=['Total Spend', 'Total Repayment', 'Due Amount', 'Profit', 'Net Payable Amount'], aggfunc='sum')

# Write the pivot table to an Excel sheet
with pd.ExcelWriter('Net_payable_amount_after_interest.xlsx') as writer:
  pivot_table.to_excel(writer, sheet_name='Output')


print("Pivot table of the analysis is written to 'spend_repayment_analysis.xlsx' file in 'Output' sheet")
