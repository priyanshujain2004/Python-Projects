import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV file into a DataFrame
df = pd.read_csv('Credit card transactions - Project - 2.csv')

# Display the first 5 rows
print(df.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df.info())

# group by `City` and sum up `Amount` field
df_result=df.groupby(['City'])['Amount'].sum().reset_index()

# calculate percentage contribution of each city
total_spend = df_result['Amount'].sum()
df_result['Percentage_Contribution'] = ((df_result['Amount'] / total_spend) * 100).round(2)

# sort in descending order of `Amount`
df_result = df_result.sort_values(by='Amount',ascending=False)

# pick top 5 cities
df_result = df_result.head(5)

# write to excel
df_result.to_excel("(task1)city_spends.xlsx", index=False)

# print the output
print(df_result.to_markdown(index=False, numalign="left", stralign="left"))
