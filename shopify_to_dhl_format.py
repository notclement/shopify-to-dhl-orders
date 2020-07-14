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
# from openpyxl import Workbook
from datetime import datetime
from shutil import copyfile

DHL_DELIMITER = '\t'
SHOPIFY_DELIMITER = ','
DHL_TEMPLATE_FILE = './resources/dhl_csv_headers.xlsx'


# EXPORT_NAME = datetime.now().strftime('export_dhl_format_%Y-%m-%d_%H%Mhr.xlsx')

def import_shopify_csv(path):
    with open(path) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=SHOPIFY_DELIMITER)
        for row in csv_reader:
            print(row)


def get_dhl_csv_format(path):
    pass


def copy_dhl_template_to_export_folder(src, dst):
    copyfile(src, dst)


def main():
    """This function will serve as a starting point for the program"""
    try:
        # shopify_csv = sys.argv[1]
        shopify_csv = './resources/shopify_csv_headers.csv'
        # import_shopify_csv(shopify_csv)

        # get_dhl_csv_format(DHL_TEMPLATE_FILE)

        dhl_export_file_name = datetime.now().strftime('export_dhl_format_%Y-%m-%d_%H%M%S.xlsx')
        copy_dhl_template_to_export_folder(DHL_TEMPLATE_FILE, './output/{}'.format(dhl_export_file_name))

    except IndexError:
        print('Please enter Shopify order csv as parameter.')


if __name__ == '__main__':
    main()
