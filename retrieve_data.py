# Retrieves data from the customer database
# pick customers added 30 days ago.
# retrieve customers added 15 days ago
# retrieve customers added 7 days ago.
# write retrieved data to a separate spreadsheet
# sync new spreadsheet to google sheets
# update sheet names on each spreadsheet written to.
import openpyxl
import config
import pdb

def main():
    thirty_days()

def get_spreadsheet_data():
    wb = openpyxl.load_workbook('customer_data-edited-dates.xlsx')
    sheet = wb.active
    # To get rows from worksheet
    rows = [row for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row)]
    return rows

def string_to_index(col):
    '''Returns Cell column index from string'''
    return openpyxl.utils.column_index_from_string(col)

def seven_days():
    pass


def fifteen_days():
    pass


def thirty_days():
    thirty_day_ws = openpyxl.Workbook()
    sheet = thirty_day_ws.active

    thirty_days_ago = config.thirty_days_ago
    rows = get_spreadsheet_data()
    # rows is a list of tuples
    # each row looks like (<Cell Sheet.A2>, <Cell Sheet.B2>, <Cell Sheet.C2>)
    retrieved_data = []

    # Search for records added in the last 30 days
    for row in rows:
        if row[4].value >= thirty_days_ago:
            retrieved_data.append(row)
    try:
        for rownum in range(len(retrieved_data) + 1):

            for row_ in retrieved_data:
                for item in row_:
                    sheet.cell(row=rownum + 1,column=string_to_index(item.column)).value \
                        = retrieved_data[rownum][string_to_index(item.column) -1].value
    except IndexError:
        pass

    thirty_day_ws.save('cust-30-days.xlsx')


    # picks customers entered within 30 days
    # data in workbook is now in rows and columns. Need to make it dict
    # https://stackoverflow.com/questions/6900955/python-convert-list-to-dictionary


if __name__ == '__main__':
    main()
