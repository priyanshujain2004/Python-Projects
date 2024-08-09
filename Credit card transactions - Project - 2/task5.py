import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV file into a DataFrame
df = pd.read_csv('Credit card transactions - Project - 2.csv')

# Display the first 5 rows
print(df.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df.info())

# Group by `City` and `Exp Type` and find the sum of `Amount`
df_result=df.groupby(['City', 'Exp Type'])['Amount'].sum()

# Sort grouped dataframe in descending order of `Amount`
df_result = df_result.sort_values(ascending=False)

# Group data on `City` pick top row's `Exp Type` and store it as `highest_expense_type`
# Group data on `City` pick bottom row's `Exp Type` and store it as `lowest_expense_type`
df_result=df_result.reset_index(level=1).groupby(level=0).agg(highest_expense_type=('Exp Type','first'), lowest_expense_type=('Exp Type','last'))

# Reset the index of the dataframe and print the columns `City`, `highest_expense_type` and `lowest_expense_type`
df_result = df_result.reset_index()
print(df_result[['City', 'highest_expense_type', 'lowest_expense_type']].to_markdown(index=False))

# Export the DataFrame to an Excel file
df_result.to_excel('(task5)city_expense_analysis.xlsx', index=False)

print("DataFrame exported to 'city_expense_analysis.xlsx' successfully!")