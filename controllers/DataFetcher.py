from splinter import Browser
from splinter import exceptions
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_col_to_name
import sys


class DataFetcher:

    def fetchData(self, inputFilename, outputFilename):
        print("Preparing Spreadsheet...")

        try:
            file = open(inputFilename, "r")
        except FileNotFoundError:
            print("No file named '" + inputFilename + "' exists in folder 'input'")
            sys.exit()

        workbook = xlsxwriter.Workbook("output/" + outputFilename + '.xlsx')
        worksheets = [workbook.add_worksheet('Computer'), workbook.add_worksheet('Biology')]

        x = 0
        row = 1
        col = 0

        print("Initiating Automation...\n")
        browser = Browser('firefox')
        browser.visit('http://result.biselahore.com/')
        browser.fill('rollNum', file.readline())
        browser.find_by_value("View Result").click()
        file.seek(0)

        for worksheet in worksheets:
            worksheet.write(0, 0, "Roll No")
            worksheet.write(0, 1, "Name")
            worksheet.set_column(xl_col_to_name(1) + ":" + xl_col_to_name(1), 24)

        row_comp = 1
        row_bio = 1

        for i in range(7):
            word = browser.find_by_tag('td')[31 + i * 9].text
            for worksheet in worksheets:
                worksheet.write(0, i + 2, word)

        worksheets[0].write(0, 9, "COMPUTER")
        worksheets[1].write(0, 9, "BIOLOGY")

        for worksheet in worksheets:
            worksheet.write(0, 10, "Total")
        browser.back()

        print("Fetching Result Data...\n")

        for line in file:
            browser.fill('rollNum', line)
            browser.find_by_value("View Result").click()
            print("Rollno: " + line)

            retry = True
            while retry:
                try:
                    elective_sub = browser.find_by_tag('td')[31 + 7 * 9].text
                    retry = False
                except IndexError:
                    retry = True
                except exceptions.ElementDoesNotExist:
                    retry = True

            if 'COMPUTER SCIENCE'.lower() in elective_sub.lower():
                row = row_comp
                worksheet = worksheets[0]
            else:
                worksheet = worksheets[1]
                row = row_bio

            worksheet.write(row, col, str(line))
            col = col + 1

            retry = True
            while retry:
                try:
                    name = browser.find_by_tag('td')[11].text
                    print(name)
                    worksheet.write(row, col, name)
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
                        marks = browser.find_by_tag('td')[38 + (index * 9)].text
                        try:
                            worksheet.write(row, col, int(marks))
                        except ValueError:
                            marks = 0
                            for i in range(3):
                                try:
                                    marks = marks + int(browser.find_by_tag('td')[i + 35 + (index * 9)].text)
                                except ValueError:
                                    continue
                            worksheet.write(row, col, marks)
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

            if 'COMPUTER SCIENCE'.lower() in elective_sub.lower():
                row_comp = row
            else:
                row_bio = row

            browser.back()

        workbook.close()
        browser.quit()
