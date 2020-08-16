# Mission-Planning-Automation-Project

OVERVIEW: Program to export current weather forecast of airports using Aviationweather.gov.  Creates spreadsheet indicating which airport 
locations and upcoming times have favorable or unfavorable weather conditions based on cloud coverage and flight visibility guidelines.

Program uses the following packages to operate; datetime, requests, pandas, re, os, tkinter, ctypes, BeautifulSoup

1. Weather rules for condition output in spreadsheet are stored in custom function.
2. Tkinter GUI to insert airport codes to be used for web scrapping feature on the Aviation website.
3. BeautifulSoup is utilized to scrap data from Request pull of HTML 
4. Pandas dataframe is used to build data table
5. Table is exported to excel and applies conditional formatting based off Weather rules function

Note: Program fails if file with same date and name title is already open.  User is prompted with this error message to close the file that is open
