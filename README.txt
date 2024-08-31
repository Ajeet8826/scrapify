Scrapify - Company Data Scraper
Scrapify is a simple desktop application designed to scrape company information from the UK Company Information service. With just a few clicks, you can extract detailed company data based on provided web links.

Version - Scrapify v1.0

Features
•	Extracts key information like company number, company name, registration date, and more.
•	Supports bulk URL processing from an Excel file.
•	Provides error handling for invalid links and incorrect file formats.
•	Generates a detailed output Excel file with three sheets:
o	Scraped: Contains successfully scraped data.
o	Invalid Links: Lists invalid or unprocessed web links.
o	Erroneous Links: Lists web links that caused errors during processing.


Installation

Run Scrapify.exe from the downloaded folder.

Usage Instructions

Prepare the Input File:

Create an Excel file named Scrapify Temp.xlsx. In this file, prepare two columns with headers:
•	Company Number
•	Web Link

Copy and Paste Data:
•	Copy and paste only the two columns—Company Number and Web Link—from your source file to the Scrapify Temp.xlsx file.

Run the Application:
•	Double-click Scrapify.exe to start the application.

Load the File:
•	When the application prompts, click Load Raw File.
•	Select your prepared Scrapify Temp.xlsx file.

Processing:
•	The application will start processing the data. The button text will change to Processing Your File to indicate the ongoing operation.

Completion:
•	Once the process is complete, the button text will change to Finished.
•	The application will save the output file with a timestamp in the format scraped_company_info_yyyy-mm-dd-HH-MM.xlsx.
Check Output:
•	Open the output file to review the scraped data and any invalid or erroneous links.
Error Handling
•	If you select a file with an incorrect format, a popup will display the message "Invalid File Selected".
•	For links that cause processing errors, they will be listed in the Erroneous Links sheet.
Deactivation and Exit
•	To exit the application, simply close the window.

