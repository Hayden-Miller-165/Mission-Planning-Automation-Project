#! python3
"""

Program to export current weather forecast of airports using Aviationweather.gov.  Resource creates spreadsheet indicating which airport 
locations and upcoming times have favorable or unfavorable weather conditions based on cloud coverage and flight visibility guidelines.

"""

# Imports packages needed for program
import datetime, requests, pandas as pd, re, os, tkinter as tk, ctypes
from bs4 import BeautifulSoup
from tkinter import messagebox

# Function to store weather rules for condition output in spreadsheet
def WeatherRules(weather):
    condition = ''
    for item in weather:
        if (item[:3] == 'OVC' or item[:3] == 'BKN') and (int(item[3:6]) <= 15):
            return 'Red'
        elif (item[:3] == 'OVC' or item[:3] == 'BKN') and (int(item[3:6]) <= 25):
            condition = 'Yellow'
        else:
            condition = 'Green'
    return condition

# Changes current working directory to folder for weather data pulls
os.chdir("FILE LOCATION HERE")

MessageBox = ctypes.windll.user32.MessageBoxW

# GUI to insert airport codes to airport code list
def GUI():
    root = tk.Tk()
    root.title("Insert Airport Code(s)")
    
    e = tk.Entry(root, width=50)
    e.pack()
    
    airport_list = ['KCLT', 'KCHS', 'KCAE', 'KGSP']
    
    AirportLabel = tk.Label(root, text='Current list: ' + 
                            ', '.join([code[1:4] for code in sorted(airport_list)]), 
                            wraplength=225)
    AirportLabel.pack(side='top')
    
    InsertLabel = tk.Label(root)
    
    def Insert():
        if not e.get():
            airportmsg = "Please insert an airport code."
        elif "K" + e.get().upper() not in airport_list:
            airportcode = "K" + e.get().upper()
            airportmsg = e.get().upper() + " inserted into airport list."
            airport_list.append(airportcode)
        else:
            airportmsg = e.get().upper() + " already in airport list.  Try again."
            
        e.delete(0, "end")
        AirportLabel.config(text='Current list: ' + 
                            ', '.join([code[1:4] for code in sorted(airport_list)]))
        InsertLabel.config(text=airportmsg)
        InsertLabel.pack(pady=5)

    def Clear():
        airport_list.clear()
        
        e.delete(0, "end")
        AirportLabel.config(text='Current list: ' + 
                            ', '.join([code[1:4] for code in sorted(airport_list)]))
        InsertLabel.config(text='List Cleared')
        InsertLabel.pack(pady=5)

    def Done():
        root.destroy()
        
    def Close():
        if messagebox.askokcancel('Quit', 'Do you want to exit the program?'):
            airport_list.clear()
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", Close)
    
    DoneButton = tk.Button(root, text="Done", command=Done, height=2, width=15)
    DoneButton.pack(side="bottom", pady=15)
    
    ClearButton = tk.Button(root, text="Clear All", command=Clear, height=2, width=15)
    ClearButton.pack(side="bottom", pady=15)
   
    InsertButton = tk.Button(root, text="Insert", command=Insert, height=2, width=15)
    InsertButton.pack(side="bottom", pady=15)
    
    root.mainloop()
    
    return airport_list

# List of airport codes
airports = '+'.join(GUI())

# Connects to URL with airport codes inserted
url = requests.get('https://aviationweather.gov/metar/data?ids={}'.format(airports) + \
'&format=raw&hours=0&taf=on&layout=on')

weatherdata = BeautifulSoup(url.content, 'lxml').find_all('code')

title = BeautifulSoup(url.content, 'lxml').find('strong').text

weatheritems = []

# Loops through HTML from BeautifulSoup pull and selects only forecasted rows
for item in weatherdata:
    kcode = re.search(r'\bK\w+', list(item.childGenerator())[0]).group() + ' '
    for text in item.childGenerator():
        if str(text) != '<br/>' and str(text)[0:4] == '\xa0\xa0FM':
            weatheritems.append(str(text).replace('\xa0\xa0FM', str(kcode)).rstrip())
            
results = pd.concat([pd.DataFrame([i], columns=[title]) for i in weatheritems],
           ignore_index=True)

# Creates various columns for spreadsheet
results['Airport'] = results[title].str.split().str[0].astype(str).str[1:]
results['Time'] = results[title].str.split().str[1]
results['Wind'] = results[title].str.split().str[2]
results['Visibility'] = results[title].str.split().str[3]
results['Cloud Coverage'] = results[title].str.split().str[4:]
results['Condition'] = [WeatherRules(i) for i in results['Cloud Coverage']]

file = datetime.datetime.today().strftime('%m-%d-%Y') + ' Aviation Weather Data.xlsx'

writer = pd.ExcelWriter(file, engine='xlsxwriter')

# Exports Pandas table to excel and applies conditional formatting based off Condition column
if len(airports) > 0:
    try:
        results.to_excel(writer, sheet_name='Aviation Weather Forecast', index=False)
    
        workbook  = writer.book
        worksheet = writer.sheets['Aviation Weather Forecast']
        
        redformat = workbook.add_format({'bg_color': '#FFC7CE',
                                      'font_color': '#9C0006'})
                                      
        yellowformat = workbook.add_format({'bg_color':   '#FFEB9C',
                                   'font_color': '#9C6500'})
        
        greenformat = workbook.add_format({'bg_color': '#C6EFCE',
                                       'font_color': '#006100'})
    
        red = {'name': 'Red', 'formattype': redformat}
        yellow = {'name': 'Yellow', 'formattype': yellowformat}              
        green = {'name': 'Green', 'formattype': greenformat}
        
        formatlist = [red, yellow, green]
        
        for formatitem in formatlist:
            worksheet.conditional_format("$A$1:$G$%d" % (len(results)+1),
                            {"type": "formula",
                             "criteria": '=INDIRECT("G"&ROW())="{}"'.format(formatitem['name']),
                             "format": formatitem['formattype']
                             })
            
        workbook.close()
    
        MessageBox(None, 'Aviation weather data pull complete.', 'Complete', 0)
    
    # Program fails if file with same date and name title is already open
    except PermissionError:
        MessageBox(None, '\nFile ' + file + ' is already open.', 'Error', 0)
else:
    MessageBox(None, 'You have now exited the program.', 'Exit', 0)
