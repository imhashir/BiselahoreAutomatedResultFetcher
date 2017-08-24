from splinter import Browser
from splinter import exceptions
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_col_to_name

class DataFetcher:
		
	def fetchData(self, inputFilename, outputFilename):
		print("Preparing Spreadsheet...")
		
		try:
			file = open(inputFilename, "r")
		except FileNotFoundError:
			print("No file named '" + inputFilename + "' exists in folder 'input'")
			sys.exit()
		
		workbook = xlsxwriter.Workbook("output/" + outputFilename + '.xlsx')
		worksheet = workbook.add_worksheet()

		x = 0
		row = 1
		col = 0

		print("Initiating Automation...\n")
		browser = Browser()
		browser.visit('http://result.biselahore.com/')
		browser.fill('rollNum', file.readline())
		browser.find_by_value("View Result").click()
		file.seek(0)

		worksheet.write(0, 0, "Roll No")
		worksheet.write(0, 1, "Name")
		worksheet.set_column(xl_col_to_name(1) + ":" + xl_col_to_name(1), 24)

		for i in range(8):
			worksheet.write(0, i + 2, browser.find_by_tag('td')[24 + i*4].text)
		worksheet.write(0, 10, "Total")
		browser.back()

		print("Fetching Result Data...\n")

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
			worksheet.write(row, col, "=SUM(" + xl_rowcol_to_cell(row, 2) + ":" + xl_rowcol_to_cell(row, col - 1) + ")")
			col = 0
			row = row + 1
			browser.back()

		workbook.close()
		browser.quit()
