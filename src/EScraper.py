from bs4 import BeautifulSoup
from datetime import datetime
import requests
import pandas as pd
import browser_cookie3
import time
import re
import sys
import urllib3


username = "jacobsardo"

# EDIT FILE NAME AND LOCATION HERE
filename = 'EtsyOrdersInfo'
location = '/Users/'+username+'/Documents/'

# EDIT EXCEL COLUMN NAMES HERE
cols = ["Order Number", "Tracking Number", "Customer"]
data = []

# Pulling Cookies
cookies = browser_cookie3.chrome(domain_name='.etsy.com')
headers = {
    "authority": "www.etsy.com",
    "method": "GET",
    "path": "/your/purchases/1661152536?ref=yr_purchases",
    "scheme": "https",
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "accept-encoding": "gzip, deflate, br",
    "accept-language": "en-US,en;q=0.9",
    "referer": "https://www.etsy.com/your/purchases?ref=hdr_user_menu-txs",
    "sec-fetch-dest": "document",
    "sec-fetch-mode": "navigate",
    "sec-fetch-site": "same-origin",
    "sec-fetch-user": "?1",
    "upgrade-insecure-requests": "1",
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36"
}
currentPage = requests.get('https://www.etsy.com/your/purchases?ref=hdr_user_menu-txs',
                           verify=True, headers=headers, cookies=cookies, timeout=15)
if (currentPage.status_code != 200):
    print('ERROR WITH VERY FIRST PAGE REQUEST')
    sys.exit()
currentPage = currentPage.text
currentPage = BeautifulSoup(currentPage, 'lxml')

# _________may not be needed anymore maybe remove this???____________________
# creating array of individual pages to loop through
pages = currentPage.find(id='content')
pages = pages.find("div", {"class": "pager"})
pages = pages.find_all("li")
# _____________________________________________________

# Looping through pages scraping info to df
pageNum = 1
while(currentPage):
    time.sleep(1)
    # Creating List of Receipt id#s
    tempPage = currentPage.find(id='content')
    receiptList = tempPage.find_all(
        "li", {"class": "receipt-section order clearfix"})
    receiptList += tempPage.find_all(
        "li", {"class": "receipt-section order clearfix placeholder loading"})
    for i in range(len(receiptList)):  # replacing each element of array with id#
        receiptList[i] = receiptList[i].get('data-receipt-id')

    # Looping through receipt id#s to find shipping number and order number
    for id in receiptList:
        dataPoint = []
        time.sleep(1)
        receiptPage = requests.get('https://www.etsy.com/your/purchases/' + id,
                                   verify=True, headers=headers, cookies=cookies, timeout=15)
        if (receiptPage.status_code != 200):
            print('ERROR WITH RECEIPT REQUEST')
            sys.exit()
        receiptPage = receiptPage.text
        receiptPage = BeautifulSoup(receiptPage, 'lxml')
        receiptPage = receiptPage.find(id='content')
        shipNum = receiptPage.find("a", {"target": "_blank"})
        customer = receiptPage.find("span", {"class": "name"}).string
        orderNum = receiptPage.find(
            "div", {"class": "user-thumb-content buyer-note labeled-section expandable last"}).text
        orderNum = re.search(r'#\d+US', orderNum)

        # adding results to dataFrame
        if(orderNum):
            print(orderNum.group())
            dataPoint.append(orderNum.group())
        else:
            print("NO ORDER NUM")
            dataPoint.append('NoneFound')
        if (shipNum):
            print('shipping number is ->>> ' +
                  shipNum.get('href').split('=')[1])
            dataPoint.append(shipNum.get('href').split('=')[1])
        else:
            print('NO SHIPPING NUM')
            dataPoint.append('NoneFound')
        dataPoint.append(customer)
        data.append(dataPoint)

    print("|")
    print("|")
    print("|")
    print("|")
    print("|")
    print("|")
    print("|")
    # finding next page
    pages = currentPage.find(id='content')
    pages = pages.find("div", {"class": "pager"})
    pages = pages.find_all("li")
    currentPage = None
    pageNum += 1
    for subpage in pages:
        print("looking for next page: " + str(pageNum))
        if (subpage.find("a", {"data-page": str(pageNum)})):
            # finding sourceHTML for next page
            href = subpage.find('a')
            href = href.get('href')
            print("\n href tag for next page is " + href + "\n")
            currentPage = requests.get('https://www.etsy.com' + href,
                                       verify=True, headers=headers, cookies=cookies, timeout=15)
            if (currentPage.status_code != 200):
                print('ERROR WITH REQUESTING SUBSEQUENT PAGE')
                sys.exit()
            currentPage = currentPage.text
            currentPage = BeautifulSoup(currentPage, 'lxml')

# Creating dataFrame and exporting to excel sheet
dateTimeObj = datetime.now()
filename = filename + ' ' + dateTimeObj.strftime("%d-%b-%Y (%H:%M)")
df = pd.DataFrame(data, columns=cols)
dfNoNames = df[["Order Number", "Tracking Number"]]
dfNoNames.to_excel(r'/Users/'+username+'/Documents/' +
                   filename + '.xlsx', index=False)
df.to_excel(r'/Users/'+username+'/Documents/' +
            filename + '|NAMES|' + '.xlsx', index=False)
