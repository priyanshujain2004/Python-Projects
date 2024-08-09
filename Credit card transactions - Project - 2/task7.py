import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Read the CSV file into a DataFrame
df_credit_card = pd.read_csv('Credit card transactions - Project - 2.csv')

# Display the first 5 rows
print(df_credit_card.head().to_markdown(index=False, numalign="left", stralign="left"))

# Print the column names and their data types
print(df_credit_card.info())

# Convert `Date` column to datetime
df_credit_card['Date'] = pd.to_datetime(df_credit_card['Date'], format='%d-%b-%y')

# Extract month and year from `Date`
df_credit_card['Month'] = df_credit_card['Date'].dt.month
df_credit_card['Year'] = df_credit_card['Date'].dt.year

# Filter data for December 2013
df_dec = df_credit_card[(df_credit_card['Year'] == 2013) & (df_credit_card['Month'] == 12)].copy()

# Filter data for January 2014
df_jan = df_credit_card[(df_credit_card['Year'] == 2014) & (df_credit_card['Month'] == 1)].copy()

# Group `df_dec` on `Card Type` and `Exp Type` and find sum of `Amount`
df_dec_grouped = df_dec.groupby(['Card Type', 'Exp Type'])['Amount'].sum().reset_index()

# Group `df_jan` on `Card Type` and `Exp Type` and find sum of `Amount`
df_jan_grouped = df_jan.groupby(['Card Type', 'Exp Type'])['Amount'].sum().reset_index()

# Merge the two dataframes on `Card Type` and `Exp Type`
df_merged = pd.merge(df_dec_grouped, df_jan_grouped, on = ['Card Type', 'Exp Type'], how='inner', suffixes=('_dec', '_jan'))

# Calculate month-over-month growth
df_merged['Growth'] = ((df_merged['Amount_jan'] - df_merged['Amount_dec']) / df_merged['Amount_dec']) * 100

# Sort in descending order of growth and pick top row
df_merged.sort_values('Growth', ascending=False, inplace=True)
df_result = df_merged.head(1)

# Write to Excel
df_result.to_excel('(task7)highest_mom_growth_jan_2014.xlsx', index=False)

# Print success message
print("Data exported to '(task7)highest_mom_growth_jan_2014.xlsx' successfully!")
