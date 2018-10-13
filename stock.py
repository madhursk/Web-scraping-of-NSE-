from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from xlwt import Workbook
import re
import xlwt
import time
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
print 'Enter the stock symbol : '
stock = raw_input()

option = webdriver.ChromeOptions()
option.add_argument("--incognito")
browser = webdriver.Chrome(executable_path='/home/madhur/Desktop/chromedriver', chrome_options=option)

browser.get("https://www.nseindia.com")
search = browser.find_element_by_id("keyword")
search.send_keys(stock)

browser.get("https://www.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuote.jsp?symbol="+str(stock)+"&illiquid=0&smeFlag=0&itpFlag=0")
def find_nth(s, x, n):
    i = -1
    for _ in range(n):
        i = s.find(x, i + len(x))
        if i == -1:
            break
    return i

titles_element = browser.find_elements_by_xpath("//div[@class='leftTableData']")
titles_name = browser.find_elements_by_xpath("//div[@class='left_info']")
titles1 = []
titles = []

for x in titles_element:
    value=x.text
    value=value.encode('ascii', 'ignore')
    titles.append(value)
    
for y in titles_name:
    value1=y.text
    value1=value1.encode('ascii', 'ignore')
    titles1.append(value1)

print(titles1)

#print(titles)
string=titles[0]
l=[]
k=0
for i in xrange(0,len(string)):
    if(string[i]=='\n'):
        value=titles[0][k:i]
        l.append(value)
        k=i+1



titles11=titles1[0]
l11=[]
k=0
for i in xrange(0,len(titles11)):
    if(titles11[i]=='\n'):
        value=titles1[0][k:i]
        l11.append(value)
        k=i+1
        

while '' in l:
    l.remove('')

l=l[1:-1]
l09=l[0:10]
l10=l[10:16]
lf=l[16:]
lfn=[]
lfn.append(l10[0]),lfn.append(str(l10[1]+l10[2])),lfn.append(l10[3]),lfn.append(str(l10[4]+l10[5]))
lmod=l09+lf+lfn
l1,l2=lmod[::2],lmod[1::2]
sheet1.write(0,0,l11[0])
sheet1.write(1,0,l11[5])


for i in xrange(1,len(l1)):
    sheet1.write(0,i,l1[i-1])
    sheet1.write(1,i,l2[i-1])
sheet1.write(0,9,l1[-1])
sheet1.write(1,9,l2[-1])
wb.save('/home/madhur/Desktop/stock.xls')
