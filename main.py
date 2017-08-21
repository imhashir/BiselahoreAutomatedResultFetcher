from splinter import Browser
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
import xlsxwriter


inputFilename = input("Enter Input Filename: ")
outFilename = input("Enter Output Filename: ")

file = open("input/" + inputFilename, "r")
workbook = xlsxwriter.Workbook("output/" + outFilename + '.xlsx')
worksheet = workbook.add_worksheet()

subCount = 8
x = 0

worksheet.write(0, 0, "RollNo")
worksheet.write(0, 1, "URDU")
worksheet.write(0, 2, "ENGLISH")
worksheet.write(0, 3, "ISLAMIYAT")
worksheet.write(0, 4, "PAK STUDIES")
worksheet.write(0, 5, "MATHS")
worksheet.write(0, 6, "PHYSICS")
worksheet.write(0, 7, "CHEMISTRY")
worksheet.write(0, 8, "COMPUTER")

row = 1
col = 0

browser = Browser()
browser.visit('http://www.biselahore.com/')
for line in file:
	browser.fill('student_rno', line)
	browser.find_by_name('submit').click()
	#TODO: Wait here
	browser.windows.current = browser.windows[1]
	while(browser.status_code != 200):
		x = x
	print("Rollno" + line)
	worksheet.write(row, col, line)
	col = col + 1
	for index in range(8):
		worksheet.write(row, col, browser.find_by_tag('td')[24 + (index*3)].text)
		col = col + 1
		#print(browser.find_by_tag('td')[23 + (index*3)].text + "\n")
		#print(browser.find_by_tag('td')[24 + (index*3)].text + "\n")
	col = 0
	row = row + 1
	print("\n\n")
	browser.windows.current = browser.windows[0]
	browser.windows[1].close()

workbook.close()
browser.quit()
