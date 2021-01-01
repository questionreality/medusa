from selenium import webdriver
from time import time
import urllib.request
import urllib.error
from urllib.request import Request, urlopen
import xlsxwriter

internet_company = "Act Fibernet"
hdr = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) ' 
		'AppleWebKit/537.11 (KHTML, like Gecko) '
		   'Chrome/23.0.1271.64 Safari/537.11',
				'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
					'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
						'Accept-Encoding': 'none',
						'Accept-Language': 'en-US,en;q=0.8',
						'Connection': 'keep-alive'}

workbook = xlsxwriter.Workbook('Task1.xlsx')

worksheet1 = workbook.add_worksheet("Sheet 1")
worksheet1.write('A1', 'WEBSITE')
worksheet1.write('B1', "LINK")
worksheet1.write('C1', "MOBILE NETWORK OPERATOR")
worksheet1.write('D1', "LINK LOAD TIME")
worksheet1.write('E1', "LINK IS DEAD OR TIMED OUT")
worksheet1_row = 1

worksheet2 = workbook.add_worksheet("Sheet 2")
worksheet2.write('A1', 'WEBSITE')
worksheet2.write('B1', "MOBILE NETWORK OPERATOR")
worksheet2.write('C1', "AVERAGE LINK LOAD TIME")
worksheet2.write('D1', "NUMBER OF DEAD LINKS")
worksheet2.write('E1', "NUMBER OF WORKING LINKS")
worksheet2.write('F1', "WEBSITE SCORE")
worksheet2_row = 1

driver = webdriver.Chrome()

web_list = ["https://nrega.nic.in/netnrega/home.aspx", 
			"https://www.usa.gov/",
			"https://www.bits-pilani.ac.in/",
			"https://www.isro.gov.in/",
			"https://medium.com/"]

web_list_number = len(web_list)
website_score_calculation_helper_list = []

for website in web_list:
	total_links_on_webpage = 0
	working_links = 0
	total_link_load_time = 0
	driver.get(website)
	all_links_list = driver.find_elements_by_xpath(".//a")
	all_links_list = [x.get_attribute("href") for x in all_links_list]

	for link_href in all_links_list:
			total_links_on_webpage+=1
			worksheet1.write(worksheet1_row, 0, website)
			worksheet1.write(worksheet1_row, 1, link_href)
			worksheet1.write(worksheet1_row, 2, internet_company)
			try:
				kick_href = ""
				page_status = ""
				page_load_time = 0
				start_time = 0
				end_time = 0
				o = ""

				for _ in range(5):
					start_time = time()
					req = Request(link_href, headers=hdr)
					page = urlopen(req)
					o = page.read()
					end_time = time()
					page_load_time+=(end_time - start_time)
					page_status = page.status
				page_load_time = page_load_time/5

				print(website, link_href, internet_company, page_load_time, "N")

				working_links+=1
				total_link_load_time+=page_load_time

				page_load_time = round(page_load_time, 5)
				worksheet1.write(worksheet1_row, 3, page_load_time)
				worksheet1.write(worksheet1_row, 4, "N")
			except urllib.error.HTTPError as e:
				hjkl = "Y, HTTPError"
				print(website, link_href, internet_company, "NaN", hjkl)
				worksheet1.write(worksheet1_row, 3, "NaN")
				worksheet1.write(worksheet1_row, 4, hjkl)
			except urllib.error.URLError as e:
				hjkl = "Y, URLError"
				print(website, link_href, internet_company, "NaN", hjkl)
				worksheet1.write(worksheet1_row, 3, "NaN")
				worksheet1.write(worksheet1_row, 4, hjkl)
			except Exception as e:
				hjkl = "Y, FATAL ERROR"
				print(website, link_href, internet_company, "NaN", hjkl)
				worksheet1.write(worksheet1_row, 3, "NaN")
				worksheet1.write(worksheet1_row, 4, hjkl)

			worksheet1_row+=1

	if working_links == 0:
		average_link_load_time = "NaN"
	else:
		average_link_load_time = total_link_load_time/working_links
		non_working_links = total_links_on_webpage - working_links

	website_score_calculation_helper_list.append([website, internet_company, average_link_load_time, non_working_links, working_links])

min_t = min([x[2] for x in website_score_calculation_helper_list])
max_t = max([x[2] for x in website_score_calculation_helper_list])

if max_t == min_t:
	min_t = 0
	max_t = 1

for i in range(0, web_list_number):
	if website_score_calculation_helper_list[i][2]=="NaN":
		website_score = -1
	else:
		A = (website_score_calculation_helper_list[i][2] - min_t)/(max_t - min_t)
		B = (website_score_calculation_helper_list[i][3])/(website_score_calculation_helper_list[i][3] + website_score_calculation_helper_list[i][4])
		website_score = (A+B)/2

	website_score_calculation_helper_list[i].append(website_score)

website_score_calculation_helper_list = sorted(website_score_calculation_helper_list, key=lambda x: x[5])

worksheet2_row = 1
for i in range(0, web_list_number):
	for j in range(0, 6):
		worksheet2.write(worksheet2_row, j, website_score_calculation_helper_list[i][j])
	worksheet2_row+=1
	

workbook.close()

