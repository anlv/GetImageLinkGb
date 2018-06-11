#!/usr/bin/python3
import sys
from openpyxl import Workbook
from openpyxl import load_workbook
import requests
import re
from bs4 import BeautifulSoup
from openpyxl.formula import Tokenizer

#read file excel using input
wb = load_workbook(sys.argv[1])
sheet = wb.worksheets[0]

#xcolumn la cot chua link trang gear
rCode = 1
rLinkImage = 4
rTitleC = 9

wCode = 1
wCodeTitle = 2
wDescription = 3
wType = 5
wTags =  6
wSKU = 14
wInventoryPolicy  = 18
wFulfillment = 19
wPrice = 20
wPriceCompare = 21
wRequireShipping = 22
wMainImage = 25
wGiftCard = 27 # value FALSE
wSubImage = 43
wWeightUnit = 44 #size (mug: oz)




xfrom = int(input ('from: '))
xto = int(input ('to: '))

# read from row to row

    # extract code

    # check config file

    # copy to link












wb.save(sys.argv[1])
