# BiselahoreAutomatedResultFetcher

BiselahoreAutomatedResultFetcher is an Automated tool that takes as input the list of Roll Numbers and then starts firefox broswer and fetches result against each Roll Number and stores subject wise marks in an Excel Sheet.

![N|Solid](http://i.imgur.com/1ahSYVK.png)

The tool was created with for personal use of automating this task but then after further improvements now it is open source for everyone.

# Details
This is a Python based tool tested and working on Python 3.6 with following modules used in creation of this tool:
  - Splinter (for browser Automation)
  - Xlsxwriter (For Excel Sheet IO)
  - wxPython (For GUI)

# How does it work
The tool reads roll numbers from file and the put each roll no turn by turn in Input field of Bise Lahore website. 
Then it extracts data from html of page and stores it in Excel Sheet.
The location of marks fields are hardcoded as their is no other way (or any way I could think of) to acheive this task.
The tool won't work if Biselahore decides to change the layout of result card or the result.biselahore.com page. 
