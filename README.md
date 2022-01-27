# webtest

This is a code that automates the process of searching for a specific term on Google. After clicking on all the relevant links, we can web scrape the required data from all the open tabs. The results are then stored in an Excel file.


This code is written in Python 3 and makes use of the following dependencies:
--> Openpyxl: To read and write into the Excel File
--> Selenium: To automate web searches
--> Requests: To get the HTML code from URLs
--> BeautifulSoup4: To get specific data from the HTML code (Web Scraping)

The code operates in the following manner:
1) Read the names of the colleges from an Excel file and store it in a list
2) Customize the query and search for it on Google
3) Navigate to the right webpage (Human)
4) Collect all the URLs and store it in a list
5) Get the HTML code from the collected URLs
6) Extract the right data from the HTML code
7) Write the results back into the Excel file

Basic error handling for the most common errors that may occur has been covered. There is further scope for error handling here.
