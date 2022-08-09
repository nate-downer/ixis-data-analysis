# Import Libraries
import numpy as np
import pandas as pd

from datetime import datetime, timedelta

# Load adds to cart data
df_adds = pd.read_csv('provided-reasources/DataAnalyst_Ecom_data_addsToCart.csv')

# Load session data
df_ses = pd.read_csv('provided-reasources/DataAnalyst_Ecom_data_sessionCounts.csv')

# Extract date time, and format month names
for i in df_ses.index:
    date = df_ses.loc[i,'dim_date']
    dt = datetime.strptime(date, "%m/%d/%y")   
    df_ses.loc[i,'dt'] = dt
    df_ses.loc[i,'month'] = str(dt.year) + '-' + str(dt.strftime('%m'))

# Create a list of "suspect browsers" containing browsers that were never actualy used to make a transaction
    # Note: These browsers are likely bots or webscrappers, and should not be counted as user sessions
    # Identifying and removing these sessions will make the data more useful to the client
browsers = df_ses[['dim_browser','sessions', 'transactions','QTY']].groupby('dim_browser').sum()
browsers.sort_values('sessions', ascending=False, inplace=True)
suspect_browsers = list(browsers[(browsers['transactions'] == 0) & (browsers['QTY'] == 0)].index)

# Use suspect_browsers list to get a cleaner set of data with most bots and web scrappers removed
clean = df_ses.drop(df_ses[df_ses['dim_browser'].isin(suspect_browsers)].index)
clean.rename(columns={'sessions':'user_sessions'}, inplace=True)
df_ses.rename(columns={'sessions':'all_sessions'}, inplace=True)

# Create Sheet 1: 
# Group the data by month and device category
sheet1 = df_ses[['dim_deviceCategory', 'transactions', 'QTY', 'month', 'all_sessions']].groupby(by=['month', 'dim_deviceCategory']).sum()

# Add a column containing the session data from non-suspect browsers
    # Note: Because of the method used to identify suspect browsers, 
    # removing them only effects session data, so alternate values for
    # transactions and quantity don't need to be included
user_sessions = clean[['dim_deviceCategory', 'user_sessions', 'month']].groupby(by=['month', 'dim_deviceCategory']).sum()
sheet1 = sheet1.join(user_sessions)

# Add an ECR column
sheet1['ECR'] = sheet1['transactions'] / sheet1['user_sessions']

# Create Sheet 2: 
# Group the data by month
sheet2 = df_ses[['transactions', 'QTY', 'month', 'all_sessions']].groupby(by=['month']).sum()

# Add a column containing the session data from non-suspect browsers
user_sessions = clean[['user_sessions', 'month']].groupby(by=['month']).sum()
sheet2 = sheet2.join(user_sessions)

# Get the data concerning the two most recent months
sheet2.sort_index(ascending=False, inplace=True)
sheet2 = sheet2.head(2)

# Create a column for ECR
sheet2['ECR'] = sheet2['transactions'] / sheet2['user_sessions']

# Add the data from df_adds
sheet2.reset_index(inplace=True)

for i in sheet2.index:
    date = sheet2.loc[i,'month'].split('-')
    year = int(date[0])
    month = int(date[1])
    x = df_adds[(df_adds['dim_year'] == year) & (df_adds['dim_month'] == month)]['addsToCart'].to_list()[0]
    sheet2.loc[i,'addsToCart'] = x

# Format the sheet to display month names and differneces
sheet2 = sheet2.transpose()

month0 = sheet2.loc['month',0]
month1 = sheet2.loc['month',1]
sheet2.rename(columns = {0:month0, 1:month1}, inplace = True)

sheet2.drop(['month'], inplace=True)
sheet2['absolute_dif'] = sheet2[month0] - sheet2[month1]
sheet2['relative_dif'] = sheet2['absolute_dif'] / sheet2[month1]

# Sheets 3 and 4 are additional analysis beyond the original scope of the assignment

# Create Sheet 3: 
# Group the session data by day of the week
sheet3 = df_ses[['dim_deviceCategory', 'transactions', 'QTY', 'all_sessions', 'dt']].copy()
sheet3 = sheet3.join(clean['user_sessions'])

sheet3['weekday_no'] = pd.DatetimeIndex(sheet3['dt']).weekday
sheet3 = sheet3[['transactions', 'QTY', 'all_sessions', 'user_sessions', 'weekday_no']].groupby('weekday_no').mean()

# Add ECR column
sheet3['ECR'] = sheet3['transactions'] / sheet3['user_sessions']

# Reformat the days of the week for readability
weekdays = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']

for i in sheet3.index:
    sheet3.loc[i,'weekday'] = weekdays[i]

sheet3.set_index('weekday', inplace=True)

# Create Sheet 4: 
# Group the session data by browser
    # Note: Because browsers were used to distinguish between 'suspect sessions' and
    # user sessions, that distinction has no meaning when the data is grouped by browser
sheet4 = df_ses[['dim_browser','all_sessions','transactions','QTY']].groupby('dim_browser').sum()

# Add ECR column
sheet4['ECR'] = sheet4['transactions'] / sheet4['all_sessions']

# Add a column indicating which browsers are included in user_sessions data
for i in sheet4.index:
    qty_val = sheet4.loc[i,'QTY']
    transactions_val = sheet4.loc[i,'transactions']

    if (qty_val == 0) & (transactions_val == 0):
        sheet4.loc[i,'browser_in_user_sessions'] = False
    else:
        sheet4.loc[i,'browser_in_user_sessions'] = True

sheet4.sort_values('all_sessions', ascending=False, inplace=True)

# Export data to Excel
with pd.ExcelWriter('python_worksheets.xlsx') as writer:
    sheet1.to_excel(writer, sheet_name='Volume by Month + Device')
    sheet2.to_excel(writer, sheet_name='Month Over Month Comparison')
    sheet3.to_excel(writer, sheet_name='Ave Volume By Weekday')
    sheet4.to_excel(writer, sheet_name='Volume by Browser')