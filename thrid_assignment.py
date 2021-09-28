import pandas as pd

df = pd.read_excel('Sales_Records_upd.xlsx')

print(len(df))

# create new columns with default value as 0
df['Margin'] = 0

df.loc[(df['Total Profit'] <= 100), 'Margin'] = "Too Low"
df.loc[(df['Total Profit'] > 100) & (df['Total Profit'] <= 5000), 'Margin'] = "Low"
df.loc[(df['Total Profit'] > 5000) & (df['Total Profit'] <= 40000), 'Margin'] = "Medium"
df.loc[(df['Total Profit'] > 40000), 'Margin'] = "High"

# show columns
print(df.columns)
# show top 5 rows
print(df.head(5))

# getting unique dates
unique_dates = df["Order Date"].unique()

# Show top 5 unique dates
print(unique_dates)

# Copying the first given unique date to a new dataframe to transfer it to excel
df1 = df.loc[(df['Order Date'] == unique_dates[1])]

# writing to excel sheet
# reading the online sales for the given date
df2 = df1.loc[(df['Sales Channel'] == 'Online')]


# region specific sum of unit sold (US) and total profits(TP) assuming US == unit cost
df['Regional sum of US and TP'] = 0
df.loc[(df['Region'] == 'Central America and the Caribbean'), 'Regional sum of US and TP'] = df['Units Sold'] + df['Total Profit']
print(df.head(5))

# Country specific sum of unit sold (US) and total profits(TP) assuming US == unit sold written in new excel file
df['Country sum of US and TP'] = 0
df.loc[(df['Country'] == 'Panama'), 'Country sum of US and TP'] = df['Units Sold'] + df['Total Profit']
df3 = df.loc[(df['Country'] == 'Panama'), 'Country sum of US and TP']

# write in multiple separate sheets
with pd.ExcelWriter('Given date sales record.xlsx') as writer:
    df1.to_excel(writer, sheet_name='Given date')
    df2.to_excel(writer, sheet_name='Online sales on given date')
df3.to_excel('Given country sales data.xlsx', index=False, sheet_name='Sales data')
# Total Revenue by region for the given year, print the output

df['Regional for a year total revenue'] = 0
df.loc[(df['Region'] == 'Asia') & (pd.DatetimeIndex(df['Order Date']).year == 2015), 'Regional for a year total revenue'] = df['Total Revenue']

print(df.head(10))







