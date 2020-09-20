from splinter import Browser
from splinter import exceptions
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_col_to_name
import sys
from time import sleep

SITE_URL = 'http://biselahore.com/SSC_A20.html'

class DataFetcher:

    def fetchData(self, inputFilename, outputFilename):
        print("Preparing Spreadsheet...")

        try:
            file = open(inputFilename, "r")
        except FileNotFoundError:
            print("No file named '" + inputFilename + "' exists in folder 'input'")
            sys.exit()

        workbook = xlsxwriter.Workbook("output/" + outputFilename + '.xlsx')
        worksheets = [workbook.add_worksheet('Computer'), workbook.add_worksheet('Biology'), workbook.add_worksheet('Other')]

        col = 0

        print("Initiating Automation...\n")
        browser = Browser('chrome')
        browser.visit(SITE_URL)
        browser.fill('student_rno', file.readline())
        sleep(0.5)
        # browser.find_by_xpath('//*[@id="main-wrapper"]/div[2]/ul/li/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td/input').click()
        file.seek(0)

        for worksheet in worksheets:
            worksheet.write(0, 0, "Roll No")
            worksheet.write(0, 1, "Name")
            worksheet.set_column(xl_col_to_name(1) + ":" + xl_col_to_name(1), 24)

        row_comp = 1
        row_bio = 1
        row_other = 1
        for i in range(7):
            word = browser.find_by_xpath('/html/body/div/ul/li/table[1]/tbody/tr[5]/td/table/tbody/tr[{}]/td[1]'.format(2 + i)).text
            for worksheet in worksheets:
                worksheet.write(0, i + 2, word)

        worksheets[0].write(0, 9, "COMPUTER")
        worksheets[1].write(0, 9, "BIOLOGY")
        worksheets[1].write(0, 9, "OTHER")

        for worksheet in worksheets:
            worksheet.write(0, 10, "Total")
        browser.back()

        print("Fetching Result Data...\n")

        for line in file:
            browser.fill('student_rno', line.strip())

            browser.find_by_xpath(
                '//*[@id="main-wrapper"]/div[2]/ul/li/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td/input').click()
            print("Rollno: " + line)

            retry = True
            while retry:
                try:
                    elective_sub = browser.find_by_xpath('/html/body/div[1]/ul/li/table[1]/tbody/tr[5]/td/table/tbody/tr[9]/td[1]').text
                    retry = False
                except IndexError:
                    retry = True
                except exceptions.ElementDoesNotExist:
                    retry = True

            if 'COMPUTER SCIENCE'.lower() in elective_sub.lower():
                row = row_comp
                worksheet = worksheets[0]
            elif 'BIOLOGY'.lower() in elective_sub.lower():
                worksheet = worksheets[1]
                row = row_bio
            else:
                worksheet = worksheets[2]
                row = row_other

            worksheet.write(row, col, str(line))
            col = col + 1
            retry = True
            while retry:
                try:
                    name = browser.find_by_xpath('/html/body/div[1]/ul/li/table[1]/tbody/tr[3]/td/table/tbody/tr[2]/td[3]').text
                    print(name)
                    worksheet.write(row, col, name)
                    retry = False
                except IndexError:
                    retry = True
                except exceptions.ElementDoesNotExist:
                    retry = True
            # /html/body/div[1]/ul/li/table[1]/tbody/tr[5]/td/table/tbody/tr[2]/td[2]
            # /html/body/div[1]/ul/li/table[1]/tbody/tr[5]/td/table/tbody/tr[3]/td[2]
            col = col + 1
            for index in range(8):
                retry = True
                while retry:
                    try:
                        marks = browser.find_by_xpath('/html/body/div[1]/ul/li/table[1]/tbody/tr[5]/td/table/tbody/tr[{}]/td[5]'.format(2 + index)).text
                        worksheet.write(row, col, int(marks))
                    except IndexError:
                        retry = True
                    except ValueError:
                        marks = 0
                        worksheet.write(row, col, int(marks))
                        retry = False
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
