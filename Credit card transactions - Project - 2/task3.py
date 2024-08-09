import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV file into a DataFrame
df = pd.read_csv('Credit card transactions - Project - 2.csv')

# Display the first 5 rows
print(df.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df.info())

# Calculate cumulative sum of `Amount` for each `Card Type`
df['Cumulative Amount'] = df.groupby('Card Type')['Amount'].cumsum()

# Filter data where `Cumulative Amount` is greater than or equal to 1000000
df_filtered = df[df['Cumulative Amount'] >= 1000000]

# Group by `Card Type` and pick the first row
df_final = df_filtered.groupby('Card Type').first().reset_index()

# Export the final data to an Excel sheet
df_final.to_excel('(task3)transaction_details.xlsx', index=False)

print("DataFrame is written to Excel File successfully.")

# print the final data
print(df_final.to_markdown(index=False, numalign="left", stralign="left"))
