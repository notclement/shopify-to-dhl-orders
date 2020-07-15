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
(./output/dhl_<export_dhl_format_YYYY-MM-DD_HHMMSS>.xlsx)
"""

import csv
from datetime import datetime
from shutil import copyfile

import openpyxl

from resources.account_details import *
from resources.dhl_index_mapping import *
from resources.shipping_service import *
from resources.shopify_index_mapping import *

DHL_DELIMITER = '\t'
SHOPIFY_DELIMITER = ','
DHL_TEMPLATE_FILE = './resources/dhl_csv_headers.xlsx'
EXPORT_NAME = datetime.now().strftime('export_dhl_format_%Y-%m-%d_%H%M%S.xlsx')
EXPORT_PATH_NAME = './output/{}'.format(EXPORT_NAME)


def get_shopify_csv_by_line(path):
    """opens and read the shopify csv and yield line by line"""
    with open(path) as csv_file:
        curr_line = 0
        csv_reader = csv.reader(csv_file, delimiter=SHOPIFY_DELIMITER)
        for row in csv_reader:
            # only yield results if its not the header
            if curr_line != 0:
                yield row
            else:
                curr_line += 1


def copy_dhl_template_to_export_folder(src, dst):
    """This function just uses the shutil copyfile func to copy the template
    DHL xlsx over to the export file to be ready to be edited"""
    copyfile(src, dst)


def convert_to_dhl(input_file_path, dhl_file_path):
    """Takes the shopify format and rearranges it into the dhl xlsx format"""
    line_num = 2
    for line in get_shopify_csv_by_line(input_file_path):
        add_line_into_dhl(line, dhl_file_path, line_num)
        line_num += 1


def add_line_into_dhl(line, dhl_file_path, line_num):
    # Call a Workbook() function of openpyxl
    # to create a new blank Workbook object
    wb = openpyxl.load_workbook(dhl_file_path)

    # Get workbook active sheet
    # from the active attribute
    ws = wb.active

    # write to the export file
    pick_up_account_number = ws[xpick_up_account_number + str(line_num)]
    pick_up_account_number.value = PICKUPACCOUNTNUMBER
    shipment_order_id = ws[xshipment_order_id + str(line_num)]
    shipment_order_id.value = line[name]
    shipping_service_code = ws[xshipping_service_code + str(line_num)]
    shipping_service_code.value = get_shipping_srv_code(line[shipping_country])
    consignee_name = ws[xconsignee_name + str(line_num)]
    consignee_name.value = line[shipping_name]
    address_line_1 = ws[xaddress_line_1 + str(line_num)]
    address_line_1.value = line[shipping_street]
    city = ws[xcity + str(line_num)]
    city.value = line[shipping_city]
    state = ws[xstate + str(line_num)]
    state.value = line[shipping_province]
    postal_code = ws[xpostal_code + str(line_num)]
    postal_code.value = line[shipping_zip]
    destination_country_code = ws[xdestination_country_code + str(line_num)]
    destination_country_code.value = line[shipping_country]
    phone_number = ws[xphone_number + str(line_num)]
    phone_number.value = line[shipping_phone]
    email_address = ws[xemail_address + str(line_num)]
    email_address.value = line[email]
    currency_code = ws[xcurrency_code + str(line_num)]
    currency_code.value = line[currency]
    total_declared_value = ws[xtotal_declared_value + str(line_num)]
    total_declared_value.value = line[total]
    shipment_description = ws[xshipment_description + str(line_num)]
    shipment_description.value = 'watch straps'
    content_description = ws[xcontent_description + str(line_num)]
    content_description.value = line[lineitem_name]
    content_unit_price = ws[xcontent_unit_price + str(line_num)]
    content_unit_price.value = line[lineitem_price]
    content_origin_country = ws[xcontent_origin_country + str(line_num)]
    content_origin_country.value = 'SG'
    content_quantity = ws[xcontent_quantity + str(line_num)]
    content_quantity.value = line[lineitem_quantity]
    content_code = ws[xcontent_code + str(line_num)]
    content_code.value = line[lineitem_sku]

    # those that need to do checks before adding
    # shipment_weight_g = ws[xshipment_weight_g+str(line_num)]
    # content_weight_g = ws[xcontent_weight_g + str(line_num)]

    # save the changes we made to the xlsx
    wb.save(dhl_file_path)


def get_shipping_srv_code(to_country):
    if to_country in lst_plt_countries:
        return 'PLT'
    else:
        return 'PPS'


def main():
    """This function will serve as a starting point for the program"""
    try:
        # shopify_csv = sys.argv[1]
        shopify_csv = './resources/shopify_csv_headers.csv'

        # cp the template into the export folder for editing
        copy_dhl_template_to_export_folder(DHL_TEMPLATE_FILE, EXPORT_PATH_NAME)

        # start the conversion process
        convert_to_dhl(shopify_csv, EXPORT_PATH_NAME)

    except IndexError:
        print('Please enter Shopify order csv as parameter.')


if __name__ == '__main__':
    main()
