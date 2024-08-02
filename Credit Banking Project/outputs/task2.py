import pandas as pd

# Read the Excel file into a DataFrame, specifying the sheet name as 'Repayment'
df_repayment = pd.read_excel('Credit Banking_Project - 1.xlsx', sheet_name='Repayment')

# Display the first 5 rows
print(df_repayment.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names of df_repayment
print(df_repayment.columns)

# Convert `Date` column to datetime
df_repayment['Date'] = pd.to_datetime(df_repayment['Date'])

# Extract year and month from `Date`
df_repayment['Year'] = df_repayment['Date'].dt.year
df_repayment['Month'] = df_repayment['Date'].dt.month_name()

# Create a dictionary mapping month names to their numerical values
month_order = {'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
               'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12}

# Map month names to numbers
df_repayment['Month_Number'] = df_repayment['Month'].map(month_order)

# Sort by month number
df_repayment = df_repayment.sort_values(by=['Year','Month_Number'])

# Group by `Customer`, `Year` and `Month` columns, sum the `Repayment` column
df_agg_repayment = df_repayment.groupby(['Customer', 'Year', 'Month_Number'])['Repayment'].sum().reset_index()


# Pivot the table to have `Year_Month` as index and `Customer` as columns and values as `Repayment`. Fill missing values with 0.
df_pivot_repayment = df_agg_repayment.pivot(index=['Year','Month_Number'], columns='Customer', values='Repayment').fillna(0)

# Write the final dataframe to an excel file called 'customer_monthly_repayment.xlsx'
df_pivot_repayment.to_excel('customer_monthly_repayment.xlsx')

