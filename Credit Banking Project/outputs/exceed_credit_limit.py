import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV files into Pandas Dataframes
df_customer_acq = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Customer Acqusition')
df_repayment = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Repayment')

# Display the first 5 rows
print("Customer Acquisition data:")
print(df_customer_acq.head().to_markdown(index=False,numalign="left", stralign="left"))

print("\nRepayment data:")
print(df_repayment.head().to_markdown(index=False,numalign="left", stralign="left"))

# Print the column names and their data types
print("\nCustomer Acquisition data:")
print(df_customer_acq.info())

print("\nRepayment data:")
print(df_repayment.info())

# Convert `Date` column to datetime
df_repayment['Date'] = pd.to_datetime(df_repayment['Date'])

# Extract year and month
df_repayment['Year'] = df_repayment['Date'].dt.year
df_repayment['Month'] = df_repayment['Date'].dt.month_name()

# Aggregate `Repayment` by `Customer`, `Year`, and `Month`
df_agg_repayment = df_repayment.groupby(['Customer', 'Year', 'Month'])['Repayment'].sum().reset_index()
df_agg_repayment = df_agg_repayment.rename(columns={'Repayment': 'Total_Repayment'})

# Merge with `df_customer_acq`
df_merged = df_agg_repayment.merge(df_customer_acq[['Customer', 'Limit']], on='Customer', how='left')

# Calculate excess spending
df_merged['Excess_Spending'] = df_merged['Total_Repayment'] - df_merged['Limit']

# Filter for excess spending
df_excess_spending = df_merged[df_merged['Excess_Spending'] > 0]

# Select and sort columns
df_final = df_excess_spending[['Year', 'Month', 'Customer', 'Excess_Spending']]
df_final = df_final.sort_values(by=['Year', 'Month'])

df_final.to_excel("exceeded_credit_limit.xlsx", index=False)

# Print the final dataframe
if not df_final.empty:
    print(df_final.to_markdown(index=False,numalign="left", stralign="left"))
else:
    print("No customers have exceeded their credit limit")
