#coding=utf-8
from bs4 import BeautifulSoup
from selenium import webdriver
from openpyxl import Workbook
import urllib
import time
import datetime

target_stock = [1303, 8422]

def lauchAndGetContent(targetId):
	targetURL = 'https://goodinfo.tw/StockInfo/StockDetail.asp?STOCK_ID='+str(targetId)
	print 'Get content from ' + targetURL
	options = webdriver.ChromeOptions()
	options.add_argument('--ignore-certificate-errors')
	driver = webdriver.Chrome(chrome_options=options)
	driver.get(targetURL)
	## Simulate the behavior of human browsing
	driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
	time.sleep(0.5)
	driver.execute_script("window.scrollTo(2, 22);")
	time.sleep(0.5)
	soup = BeautifulSoup(driver.page_source, 'html.parser')
	table = soup.find(id='SALE_MONTH_INFO')
	time.sleep(1.0)
	driver.close()
	driver.quit()
	return table

def getLatestMonth():
	now = datetime.datetime.now()
	year = str(now.year).replace("20","")
	month_digit = now.month - 1
	month = ''
	if month_digit < 10:
		month = '0'+str(month_digit)
	latest = year + '/' + month
	print 'getLatestMonth: ' + latest
	return latest

def retrieveLatestMonthlyReport(table):
	table_data = [[cell.text for cell in row("td")] for row in table.find_all('tr')]
	table_list = []
	latestMonth = getLatestMonth()
	for id, data in enumerate(table_data):
		if latestMonth in data or len(table_list) != 0:
			table_list.append(data)
	return table_list

def generateExcel(stock_dict):
	wb = Workbook()
	# grab the active worksheet
	ws = wb.active
	for stock_id, data_list in stock_dict.iteritems():
		ws = wb.create_sheet(stock_id)
		ws.append(['年/月', '單月營收億元', '單月月增%', '單月年增%', '累計營收億元', '累計年增'])
		for row in data_list:
			ws.append(row)
	wb.save("stock.xlsx")

if __name__ == "__main__":
	print 'start scraping web'
	stock_dict = {}
	for stock_id in target_stock:
		data_list = lauchAndGetContent(stock_id)
		monthly_revenue = retrieveLatestMonthlyReport(data_list)
		stock_dict[str(stock_id)] = monthly_revenue
	
	generateExcel(stock_dict)



