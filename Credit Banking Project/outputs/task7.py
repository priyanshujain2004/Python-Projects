import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the excel file into a DataFrame
df_spend = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Spend')

# Display the first 5 rows
print(df_spend.head().to_markdown(index=False,numalign="left", stralign="left"))

# Print the column names and their data types
print(df_spend.info())

# Convert `Date` to datetime
df_spend['Date'] = pd.to_datetime(df_spend['Date'], format='%Y-%m-%d')

# Extract year and month
df_spend['Year'] = df_spend['Date'].dt.year
df_spend['Month'] = df_spend['Date'].dt.month_name()

# Calculate total spend by category, year, and month
df_agg = df_spend.groupby(['Category', 'Year', 'Month'])['Spend'].sum().reset_index(name='Total_Spend')

# Sort by year, month, and total spend
df_agg_sorted = df_agg.sort_values(by=['Year', 'Month', 'Total_Spend'], ascending=[True, True, False])

# Get the top category for each month
df_top_category = df_agg_sorted.groupby(['Year', 'Month']).first().reset_index()

# Create a pivot table
pivot_table = df_top_category.pivot(index=['Year', 'Month'], columns='Category', values='Total_Spend').fillna(0)

# Print the pivot table
print("The pivot table of the total spend per category per month is as following:")
print(pivot_table.to_markdown(index=False,numalign="left", stralign="left"))

# Write the pivot table to an excel file
pivot_table.to_excel('monthly_spend_category.xlsx')
