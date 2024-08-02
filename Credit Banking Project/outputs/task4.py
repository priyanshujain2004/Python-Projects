import pandas as pd

# Read the excel files into Pandas Dataframes
df_spend = pd.read_excel('Credit Banking_Project - 1.xlsx',sheet_name='Spend')
df_customer_acq = pd.read_excel('Credit Banking_Project - 1.xlsx',sheet_name='Customer Acqusition')

# Display the first 5 rows
print("First 5 rows of Spend Data:")
print(df_spend.head().to_markdown(index=False, numalign="left", stralign="left"))
print("\nFirst 5 rows of Customer Acquisition Data:")
print(df_customer_acq.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print("\nSpend Data column names and types:")
print(df_spend.info())
print("\nCustomer Acquisition Data column names and types:")
print(df_customer_acq.info())

# Convert `Date` column to datetime
df_spend['Date'] = pd.to_datetime(df_spend['Date'])

# Extract month and year and create `Month_Year` column
df_spend['Month_Year'] = df_spend['Date'].dt.strftime('%Y-%m')

# Merge the dataframes on `Customer`
merged_df = df_spend.merge(df_customer_acq, on='Customer', how='inner')

# Group by `Segment` and `Month_Year` and calculate total `Spend`
spend_segment_month_wise = merged_df.groupby(['Segment', 'Month_Year'])['Spend'].sum().reset_index()

# Sort by `Month_Year` and `Spend` in descending order
spend_segment_month_wise_sorted = spend_segment_month_wise.sort_values(by=['Month_Year', 'Spend'], ascending=[True, False])

# Group by `Month_Year` and select the first row (highest spending segment)
highest_spending_segment_month_wise = spend_segment_month_wise_sorted.groupby('Month_Year').first().reset_index()

# Write the final dataframe to a new excel file
highest_spending_segment_month_wise.to_excel('highest_spending_segment_month_wise.xlsx', index=False)

# Display the first 5 rows
print("\nHighest spending segment each month:")
print(highest_spending_segment_month_wise.head(5).to_markdown(index=False, numalign="left", stralign="left"))
