import datetime


# Customer Information Workbook
customer_data = 'customer_data.xlsx'

# Dates
today = datetime.datetime.today()
thirty_days_ago = today - datetime.timedelta(days=30)
fifteen_days_ago = today - datetime.timedelta(days=15)
seven_days_ago = today - datetime.timedelta(days=7)
