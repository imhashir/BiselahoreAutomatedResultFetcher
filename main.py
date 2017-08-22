from splinter import Browser
from splinter import exceptions

import xlsxwriter
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait

inputFilename = input("Enter Input Filename: ")
outFilename = input("Enter Output Filename: ")

file = open("input/" + inputFilename, "r")
workbook = xlsxwriter.Workbook("output/" + outFilename + '.xlsx')
worksheet = workbook.add_worksheet()

x = 0

worksheet.write(0, 0, "RollNo")
worksheet.write(0, 1, "Name")
worksheet.write(0, 2, "URDU")
worksheet.write(0, 3, "ENGLISH")
worksheet.write(0, 4, "ISLAMIYAT")
worksheet.write(0, 5, "PAK STUDIES")
worksheet.write(0, 6, "MATHS")
worksheet.write(0, 7, "PHYSICS")
worksheet.write(0, 8, "CHEMISTRY")
worksheet.write(0, 9, "COMPUTER")

row = 1
col = 0

browser = Browser()
browser.visit('http://www.biselahore.com/')
for line in file:
	browser.fill('student_rno', line)
	browser.find_by_name('submit').click()
	browser.windows.current = browser.windows[1]
	print("Rollno" + line)
	worksheet.write(row, col, int(line))
	col = col + 1
	
	retry = True
	while retry:
		try:
			worksheet.write(row, col, browser.find_by_tag('td')[8].text)
			retry = False
		except IndexError:
			print("Retrying Name...")
			retry = True
		except exceptions.ElementDoesNotExist:
			print("Retrying Name...")
			retry = True
	
	col = col + 1
	for index in range(8):
		retry = True
		while retry:
			try:
				worksheet.write(row, col, int(browser.find_by_tag('td')[24 + (index*3)].text))
			except IndexError:
				retry = True
			except exceptions.ElementDoesNotExist:
				retry = True
			else:
				retry = False
				
		col = col + 1
	col = 0
	row = row + 1
	browser.windows.current = browser.windows[0]
	browser.windows[1].close()

workbook.close()
browser.quit()
