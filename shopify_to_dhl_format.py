"""
Created by: Clement

Shopify orders to DHL shipping format

Will be able to let me export orders via Shopify and then convert it into the
format that DHL portal can take in directly
https://ecommerceportal.dhl.com/Portal/pages/login/userlogin.xhtml

STEPS:
1. Logistics team will just export the Shopify orders as csv
2. Run the script with the shopify exported order list
3. DHL formatted csv will be in an output folder
(./output/dhl_<shopify_filename>.csv)
"""

import csv
import xlrd

DHL_DELIMITER = '\t'
SHOPIFY_DELIMITER = ','
DHL_HEADERS_FILE = './resources/dhl_csv_headers.xlsx'


def import_shopify_csv(path):
    with open(path) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=SHOPIFY_DELIMITER)
        for row in csv_reader:
            print(row)


def get_dhl_csv_format(path):
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)

    for row in range(sheet.nrows):
        values = sheet.row_values(row)
        print(values)


def main():
    """This function will serve as a starting point for the program"""
    try:
        # shopify_csv = sys.argv[1]
        shopify_csv = './resources/shopify_csv_headers.csv'
        # import_shopify_csv(shopify_csv)

        get_dhl_csv_format(DHL_HEADERS_FILE)

    except IndexError:
        print('Please enter Shopify order csv as parameter.')


if __name__ == '__main__':
    main()
