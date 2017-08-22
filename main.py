from splinter import Browser
from splinter import exceptions

import xlsxwriter

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
browser.visit('http://result.biselahore.com/')

for line in file:
	browser.fill('rollNum', line)
	browser.find_by_value("View Result").click()
	print("Rollno" + line)
	worksheet.write(row, col, str(line))
	col = col + 1
	
	retry = True
	while retry:
		try:
			worksheet.write(row, col, browser.find_by_tag('td')[11].text)
			retry = False
		except IndexError:
			retry = True
		except exceptions.ElementDoesNotExist:
			retry = True
	
	col = col + 1
	for index in range(8):
		retry = True
		while retry:
			try:
				worksheet.write(row, col, int(browser.find_by_tag('td')[26 + (index*4)].text))
			except IndexError:
				retry = True
			except exceptions.ElementDoesNotExist:
				retry = True
			else:
				retry = False
				
		col = col + 1
	col = 0
	row = row + 1
	browser.back()

workbook.close()
browser.quit()
