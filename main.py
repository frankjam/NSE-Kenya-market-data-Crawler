import requests, os
import lxml
from bs4 import BeautifulSoup
from xlwt import *
from datetime import date

#workbook to store collected date
workbook = Workbook(encoding = 'utf-8')
table = workbook.add_sheet('data')
#Create the header of each column in the first row.
table.write(0, 0, 'Number')
table.write(0, 1, 'Stock url')
table.write(0, 2, 'Stock name')
table.write(0, 3, 'Stock payout rate')
table.write(0, 4, 'Stock payout date')
line = 1

#array variables to store the data crawled
dividedname = []
dividedurl = []
divideddate = []
dividedrate = []

url = "https://live.mystocks.co.ke/"
headers = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36 QIHU 360SE'
}

f = requests.get(url, headers = headers)
soup =BeautifulSoup(f.content,'lxml')

#crawling the data fetched from the url 
dividendsinfo = soup.find("div", {"id": "hpEvents"}).find_all('a')
for anchor in dividendsinfo:
    urls = url + anchor['href']
    dividedurl.append(urls)
    dividedname.append(anchor.text)

dividendsmoreinfo = soup.find("div", {"id": "hpEvents"});
payoutdate = dividendsmoreinfo.find_all('span',{"class":'hpStamp'})
for pd in payoutdate:
    divideddate.append(pd.get_text())

payoutrate = dividendsmoreinfo.find_all('span',{"class":'hpEvent'})
for pr in payoutrate:
    payrate = pr.get_text()
    dividedrate.append(payrate[4:])

itemsnumber = len(dividedname)
for x in range(itemsnumber):
    #Write the crawled data into Excel separately from the second row.
    table.write(line, 0, x+1)
    table.write(line, 1, dividedurl[x])
    table.write(line, 2, dividedname[x])
    table.write(line, 3, dividedrate[x])
    table.write(line, 4, divideddate[x])
    line +=1

#saving the data to a workbook
timetoday =date.today().strftime("%d %b %Y")
workbook.save('stock saved on '+timetoday+'.xls')

#check if file was saved
path = os.getcwd() +'\stock saved on '+timetoday+'.xls'
if(os.path.exists(path)):
    print("Print workbook saved successfully")
else:
    print("Print workbook was not saved successfully")