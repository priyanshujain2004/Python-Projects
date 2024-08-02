import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV files into Pandas Dataframes
df_spend = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Spend')
df_repayment = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Repayment')
df_customer_acquisition = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Customer Acqusition')

# Display the first 5 rows of each DataFrame
print("First 5 rows of Spend Data:")
print(df_spend.head().to_markdown(index=False, numalign="left", stralign="left"))
print("\n")

print("First 5 rows of Repayment Data:")
print(df_repayment.head().to_markdown(index=False, numalign="left", stralign="left"))
print("\n")

print("First 5 rows of Customer Acquisition Data:")
print(df_customer_acquisition.head().to_markdown(index=False, numalign="left", stralign="left"))
print("\n")

# Print the column names and their data types of each DataFrame
print("Spend Data Info:")
print(df_spend.info())
print("\n")

print("Repayment Data Info:")
print(df_repayment.info())
print("\n")

print("Customer Acquisition Data Info:")
print(df_customer_acquisition.info())

# Convert `Date` columns to datetime format
df_spend['Date'] = pd.to_datetime(df_spend['Date'])
df_repayment['Date'] = pd.to_datetime(df_repayment['Date'])

# Extract `Year` and `Month` from `Date` column
df_spend['Year'] = df_spend['Date'].dt.year
df_spend['Month'] = df_spend['Date'].dt.month_name()
df_repayment['Year'] = df_repayment['Date'].dt.year
df_repayment['Month'] = df_repayment['Date'].dt.month_name()

# Merge `df_spend` and `df_repayment` on `Customer` and `Date` columns
df_merged = df_spend.merge(df_repayment, on=['Customer', 'Date'], how='inner')

# Calculate `Profit` as `Spend` - `Repayment` * 0.029
df_merged['Profit'] = (df_merged['Spend'] - df_merged['Repayment']) * 0.029

# Merge the result with `df_customer_acquisition` on `Customer` column
df_merged = df_merged.merge(df_customer_acquisition, on='Customer', how='inner')

# Group by `Segment`, `Year_x`, and `Month_x` and sum the `Profit`
df_agg = df_merged.groupby(['Segment', 'Year_x', 'Month_x'])['Profit'].sum().reset_index()

# Sort the grouped by dataframe in descending order of `Profit` and pick top value for each `Year_x` and `Month_x` combination
df_agg = df_agg.sort_values(by='Profit', ascending=False).groupby(['Year_x', 'Month_x']).head(1)

# Rename the columns `Year_x` to `Year` and `Month_x` to `Month`
df_agg = df_agg.rename(columns={'Year_x': 'Year', 'Month_x': 'Month'})

# Create a pivot table to get the most profitable `Segment` for each `Month` and `Year` combination
pivot_table = df_agg.pivot(index=['Year', 'Month'], columns='Segment', values='Profit')

# Save the pivot table in an excel file named "most_profitable_segment.xlsx"
pivot_table.to_excel('most_profitable_segment.xlsx')

