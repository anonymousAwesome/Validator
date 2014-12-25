#Note: this script requires the following libraries to be installed:
#
#openpyxl 
#http://openpyxl.readthedocs.org/
#
#lxml
#http://lxml.de/
#
#requests
#http://docs.python-requests.org/

import requests
import time
from lxml import html
from openpyxl import load_workbook

def pull_website(ProductID):
	
	#Defines the parameters for the GET request:
	URL = 'http://www.costco.com/CatalogSearch'
	payload = {
		'storeId':'10301',
		'catalogId': '10701',
		'langId':'-1',
		'keyword': ProductID,
		'refine':''
	}
	
	#Downloads the website
	page = requests.get(URL, params=payload)
	
	#Converts the website into a form the parser can search through
	tree = html.fromstring(page.text)
	
	#Gathers the required information from the downloaded website using XPath searches:	
	if tree.xpath('//*[@id="main_content_wrapper"]//span[@itemprop="sku"]/text()')!=[]:
		ItemNumber=int(tree.xpath('//*[@id="main_content_wrapper"]//span[@itemprop="sku"]/text()')[0])
		ProductName=unicode(tree.xpath('//div[@class="product-info"]/div[@class="top_review_panel"]/h1[@itemprop="name"]/text()')[0]).strip()
		Brand=unicode(tree.xpath('//div[@id="product-tab2"]//li/text()')[2]).strip()
		catEntryId=int(tree.xpath('//ul[@class="products"]/li[@class="product"]/input[@name="catEntryId"]/@value')[0])
		categoryId=int(tree.xpath('//*[@id="ProductForm"]//input[@name="categoryId"]/@value')[0])

		return {'ProductID': ProductID, 'ItemNumber':ItemNumber, 'ProductName': ProductName, 'Brand': Brand, 'catEntryId': catEntryId, 'categoryId': categoryId}
	else: 
		print ""
		print "Whoops, Product ID",ProductID,"couldn't be found!"
		return ""
		
#Loads the excel file and selects the active worksheet
wb=load_workbook("CostcoSamples.xlsx")
ws=wb.active

def make_dictionary_from_xlsx(row):

	return {
	'ProductID': ws.rows[row][0].value,
	'ItemNumber':ws.rows[row][1].value,
	'ProductName': ws.rows[row][2].value,
	'Brand': ws.rows[row][14].value,
	'catEntryId': ws.rows[row][33].value,
	'categoryId': ws.rows[row][35].value
	}

for row in range(1,len(ws.rows)):
	xlsx_dict=make_dictionary_from_xlsx(row)
	website_dict=pull_website(xlsx_dict["ProductID"])
	time.sleep(1)
	for keys in xlsx_dict:
		if website_dict!="":
			if xlsx_dict[keys] != website_dict[keys]:
				print ""
				print "Product ID:",xlsx_dict["ProductID"]
				print keys
				print (xlsx_dict[keys]),
				print '->',
				print (website_dict[keys])
