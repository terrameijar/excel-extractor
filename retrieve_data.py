# Retrieves data from the customer database and writes it to a new spreadsheet
import sys
import argparse
import openpyxl
import config


def main(n):
    create(n)


def get_spreadsheet_data():
    """Reads data from the original customer spreadsheet"""
    try:
        wb = openpyxl.load_workbook(config.customer_data)
        sheet = wb.active
        # To get rows from worksheet
        rows = \
            [row for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row)]
        return rows
    except IOError:
        raise IOError('Customer database not found. Ensure that the customer '
                      'info spreadsheet exists and that the config file points'
                      ' to it')


def string_to_index(col):
    """Returns Cell column index from string"""
    return openpyxl.utils.column_index_from_string(col)


def create(days):
    """Creates new spreadsheet with customer info added in the last n days

    Attributes:
        days: int. The search query will search for records added in this
             number of days.
    """

    if days == 7:
        date_delta = config.seven_days_ago
    elif days == 15:
        date_delta = config.fifteen_days_ago
    elif days == 30:
        date_delta = config.thirty_days_ago
    else:
        raise ValueError('Incorrect date range entered')

    new_cust_ws = openpyxl.Workbook()
    sheet = new_cust_ws.active
    rows = get_spreadsheet_data()
    # rows is a list of tuples
    # each row looks like (<Cell Sheet.A2>, <Cell Sheet.B2>, <Cell Sheet.C2>)
    retrieved_data = []

    # Search for records added in the last n days
    for row in rows:
        if row[4].value >= date_delta:
            retrieved_data.append(row)
    try:
        for rownum in range(len(retrieved_data) + 1):

            for row_ in retrieved_data:
                for item in row_:
                    sheet.cell(row=rownum + 1, column=string_to_index(item.column)).value \
                        = retrieved_data[rownum][string_to_index(item.column) - 1].value

    except IndexError:
        pass

    new_cust_ws.save('cust-{0}-days.xlsx'.format(days))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        prog="retrieve_data", description="Retrieve data from a spreadsheet",
        usage='%(prog)s [options]')
    parser.add_argument(
        "days", type=int, help="Retrieve customer data from the last n days",
        choices=[7, 15, 30])
    args = parser.parse_args()
    main(args.days)
