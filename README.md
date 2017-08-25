# BiselahoreAutomatedResultFetcher

BiselahoreAutomatedResultFetcher is an Automated tool that takes as input the list of Roll Numbers in plain text file, starts firefox browser and fetches result against each Roll Number and stores subject wise marks in an Excel Sheet.

![N|Solid](http://i.imgur.com/1ahSYVK.png)

The tool was created with for personal use of automating this task but then after further improvements now it is open source for everyone.

# Details
This is a Python based tool tested and working on Python 3.6 with following modules used in creation of this tool:
  - [Splinter](https://splinter.readthedocs.io/en/latest/install.html) (for browser Automation)
  - [Xlsxwriter](http://xlsxwriter.readthedocs.io/getting_started.html) (For Excel Sheet IO)
  - [wxPython](https://www.wxpython.org/pages/downloads/) (For GUI)

# How to Use
  - Firefox browser installed
  - You need Python3 installation in your machine along with all the modules mentioned in "How does it work" section.
  - Then open command windows/terminal in project directory and type:
    ```sh
    python main.py
    ```
  - Then select yout plain text input file with following format:
    ```sh
    rollno1
    rollno2
    rollno3
    ```
  - Then in next field, Enter your output filename. This will be the name of Excel file that will be generated at the end of execution with all subject wise details of studnets. This file will be generated in the folder named "Output"
  - Hit the Start button, sit back and wait for automated system to do all the task for you.

# How does it work
The tool reads roll numbers from file and the put each roll no turn by turn in Input field of Bise Lahore website. 
Then it extracts data from 'td' of result card table and stores it in Excel Sheet.
The location of marks fields are hardcoded as their is no other way (or any way I could think of) to acheive this task.
The tool won't work if Biselahore decides to change the layout of result card or the result.biselahore.com page. So read below to get it to work properly in case anything changes. 

In case the design of Result card changes, you need to modify the following lines in controllers/DataFetcher.py
```sh
#For Student Name - 11 is the 'td' number of name
worksheet.write(row, col, browser.find_by_tag('td')[11].text)
#For Subject Names - 24 is the 'td' number of 1st subject name
worksheet.write(0, i + 2, browser.find_by_tag('td')[24 + i*4].text)
#For marks - 26 is the 'td' number of 1st subject marks
worksheet.write(row, col, int(browser.find_by_tag('td')[26 + (index*4)].text)) 
```
and set appropriate number of 'td' for app to work properly.
