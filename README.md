# shopify-to-dhl-orders
 _This is to convert Shopify order export (csv) into DHL AP input format (xlsx)_

**To get started**
1. Create `account_details.py` in the resources folder
2. In that file, add `PICKUPACCOUNTNUMBER = 'xxx'` where xxx is your Pick-up Account Number provided by DHL
3. To run - `python main.py`
4. On a browser, go to `<yourip>:5000`

![landingpage](https://i.imgur.com/dFcgZRQ.png)
5. Upload a file and it will process it and send it back to you once it is done.

![file selected](https://i.imgur.com/Cs8Vln5.png)

![File returned from server](https://i.imgur.com/ntUwVTt.png)

**Misc**
1. To change your item stats, edit `.\resources\strap_stats.py`

**Dependencies**
1. openpylx - https://openpyxl.readthedocs.io/en/stable/

> # TODO
> 1. Input Validation checks