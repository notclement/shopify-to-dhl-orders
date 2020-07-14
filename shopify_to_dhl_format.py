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

import sys
from account_details import *
from strap_stats import *
from shipping_service import *


def main():
    """This function will serve as a starting point for the program"""
    print(PICKUPACCOUNTNUMBER)


if __name__ == '__main__':
    main()
