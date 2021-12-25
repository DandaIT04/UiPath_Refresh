# UiPath_Refresh
Trying to refresh usage of UiPath by attempting to redo my last assignment. Basically, i am to retrieve share prices of coca cola from MSN by scraping the data to xlsx file, create a chart based on date and open share price to be able to see share price increase over time and send the email to my teammates.

## Design Process
-  Open browser with URL set to "https://www.msn.com/en-sg/money/stockdetails/history/fi-a1wljc", to retrieve history of Coca Cola share prices
-  Scrape the data, most important columns are date and open share prices, save it to a data table.
-  Create a new folder in desktop to save extracted data table into xlsx file
-  Use excel application scope to create new xlsx file called test.xlsx and write the datatable into the excel files using write range
-  Open excel file using open application scope, navigate using mouse clicks to create a line chart of date and open share prices,sorted from oldest to new share price
-  Send excel file to receiver

