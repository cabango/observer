#coding=utf-8
#coding=big5
from bs4 import BeautifulSoup
from selenium import webdriver
from openpyxl import Workbook
from yattag import Doc
import urllib
import time
import datetime
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

target_stock = {'1303':'南亞', '8422':'可寧衛', '5871':'中租', '1215':'卜蜂', '9943':'好樂迪'}
montly_title = ['年/月', '單月營收億元', '單月月增%', '單月年增%', '累計營收億元', '累計年增']

def lauchAndGetContent(targetId):
	targetURL = 'https://goodinfo.tw/StockInfo/StockDetail.asp?STOCK_ID=' + targetId
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
	#table = soup.find(id='SALE_MONTH_INFO')
	time.sleep(1.0)
	driver.close()
	driver.quit()
	return soup

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

def generateHtmlOutput(stock_dict):
	doc, tag, text = Doc().tagtext()
	with tag('html'):
		with tag('head'):
			with tag('link', 'rel=\'stylesheet\'', 'href=\'style.css\''):
				text('')
			with tag('body'):
				for id, stock in stock_dict.iteritems():
					stock.scrapeMonthlyReport()
					report = stock.monthlyReport
					with tag('div', 'id=\'DIV2\''):
						with tag('table'):
							with tag('tr'):
								with tag('td', 'colspan=\'6\''):
									text(stock.id+stock.name)
							with tag('tr'):
								for title in montly_title:
									with tag('th'):
										text(title)
							for row in report:
								with tag('tr'):
									for column in row:
										with tag('td'):
											text(column)
	'''
	content = '<html><head><link rel=\'stylesheet\' href=\'style.css\'></link><body>'
	for id, stock in stock_dict.iteritems():
		stock.captureMonthlyReport()
		content += '<div id=\'DIV2\'>'
		content += str(stock.monthlyReport)
		content += '</div>'
	content += '</body></head></html>'
	print content
	'''

	f = open('MonthlyReport.htm', 'w')
	f.write(doc.getvalue())
	f.close()

class Stock(object):
	"""docstring for Stock"""
	def __init__(self, id, name, rawHtml):
		self.id = id
		self.name = name
		self.rawHtml = rawHtml

	def scrapeMonthlyReport(self):
		monthlyReport = self.rawHtml.find(id='SALE_MONTH_INFO')
		table_data = [[cell.text for cell in row("td")] for row in monthlyReport.find_all('tr')]
		table_list = []
		latestMonth = getLatestMonth()
		for id, data in enumerate(table_data):
			#if latestMonth in data or len(table_list) != 0:
			if latestMonth in data:
				table_list.append(data)
		self.monthlyReport = table_list

	def printInfo(self):
		print self.id
		print self.name
		print self.monthlyReport

if __name__ == "__main__":
	print 'start scraping web'
	stock_dict = {}
	for stock_id, stock_name in target_stock.iteritems():
		rawHtml = lauchAndGetContent(stock_id)
		stock = Stock(stock_id, stock_name, rawHtml)
		#monthly_revenue = retrieveLatestMonthlyReport(data_list)
		stock_dict[stock_id] = stock
	
	generateHtmlOutput(stock_dict)


