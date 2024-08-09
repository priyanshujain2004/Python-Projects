import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV file into a DataFrame
df_credit_card = pd.read_csv('Credit card transactions - Project - 2.csv')

# Display the first 5 rows
print(df_credit_card.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df_credit_card.info())

# Group data on `Exp Type` and find the sum of `Amount`
df_result = df_credit_card.groupby('Exp Type')['Amount'].sum().reset_index()

# Rename the `Amount` column to `Total Spends`
df_result = df_result.rename(columns={'Amount':'Total Spends'})

# Filter data on `Gender` = 'F' and group on `Exp Type` to find the sum of `Amount`
df_female_spend = df_credit_card[df_credit_card['Gender']=='F'].groupby('Exp Type')['Amount'].sum().reset_index()

# Rename the `Amount` column to `Female Spends`
df_female_spend = df_female_spend.rename(columns={'Amount':'Female Spends'})

# Merge the two dataframes on `Exp Type`
df_result = pd.merge(df_result, df_female_spend, on='Exp Type', how='left')

# Calculate the percentage contribution of spends by females
df_result['Percentage Contribution'] = (df_result['Female Spends']/df_result['Total Spends'])*100

# Format the result in percentage upto 2 decimal places
df_result['Percentage Contribution'] = df_result['Percentage Contribution'].map('{:,.2f}%'.format)

# Sort grouped dataframe in descending order of `Percentage Contribution`
df_result.sort_values('Female Spends', ascending=False, inplace=True)

# Write the dataframe to excel
df_result.to_excel('(task6)percentage_contribution_female_spends.xlsx', index=False)

# Print the success message
print("Data exported to '(task6)percentage_contribution_female_spends.xlsx' successfully!")
