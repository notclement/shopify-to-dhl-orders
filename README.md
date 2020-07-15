# shopify-to-dhl-orders
 _This is to convert Shopify order export (csv) into DHL AP input format (xlsx)_

**To get started**
1. Create `account_details.py` in the resources folder
2. In that file, add `PICKUPACCOUNTNUMBER = 'xxx'` where xxx is your Pick-up Account Number provided by DHL
3. To run - `python shopify_to_dhl_format.py <shopify_export.csv>`
4. Output path and format -  `./output/dhl_<export_dhl_format_YYYY-MM-DD_HHMMSS>.xlsx`

**Misc**
1. To change your item stats, edit `.\resources\strap_stats.py`

**Dependencies**
1. openpylx - https://openpyxl.readthedocs.io/en/stable/

> # TODO
> 1. Add front end functionality with Flask (basic upload file page and also a page that shows the output folder to download stuff)
> 2. Input Validation checks