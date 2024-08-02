import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the excel files into Pandas Dataframes
df_customer_acquisition = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Customer Acqusition')
df_spend = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Spend')

# Display the first 5 rows
print("Customer Acquisition Data:")
print(df_customer_acquisition.head().to_markdown(index=False, numalign="left", stralign="left"))

print("\nSpend Data:")
print(df_spend.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print("\nCustomer Acquisition Data:")
print(df_customer_acquisition.info())

print("\nSpend Data:")
print(df_spend.info())

# Convert `Date` column to datetime
df_spend['Date'] = pd.to_datetime(df_spend['Date'])


# Create age groups
bins = [18, 30, 50, float('inf')]
labels = ['18-30', '30-50', '50+']
df_customer_acquisition['Age_Group'] = pd.cut(df_customer_acquisition['Age'], bins=bins, labels=labels, right=False)

# Merge dataframes
merged_df = df_spend.merge(df_customer_acquisition[['Customer', 'Age_Group']], on='Customer', how='left')

# Extract year and month
merged_df['Year'] = merged_df['Date'].dt.year
merged_df['Month'] = merged_df['Date'].dt.month_name()

# Calculate total spend by age group, year, and month
grouped_df = merged_df.groupby(['Age_Group', 'Year', 'Month'])['Spend'].sum().reset_index()

# Sort by spend in descending order
grouped_df = grouped_df.sort_values(by='Spend', ascending=False)

# Get the highest spending age group for each month
highest_spending_df = grouped_df.groupby(['Year', 'Month']).first().reset_index()

# Create a pivot table
pivot_table = highest_spending_df.pivot(index='Month', columns='Year', values='Age_Group')

# Write to Excel
pivot_table.to_excel('highest_spending_age_group_by_month.xlsx')
