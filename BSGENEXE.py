
#!pip install tkcalendar
#!pip install tkinter
#!pip install gspread
#!pip install xlwings
#!pip install oauth2client
import tkinter as tk
from tkinter import *
from tkcalendar import Calendar
from datetime import date
import calendar
import datetime
import pandas as pd
import xlwings as xw
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
#import openpyxl as xl
#from openpyxl import Workbook
#Authorizing Credentials to Access Google Sheets API 
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
sh = client.open("Copy of Project Coverage Sheet")
def api_handler(day, area_time):
    #Creating Pandas DataFrame for all 
    worksheet = sh.worksheet('raw_dataframe')
    df = pd.DataFrame(worksheet.get_all_records())
    df = pd.concat([df['Name'], df['Break'].str.split('-', expand=True), df['Shift_Start'], df['Shift_End'], df['Day'], df['Area_Time'], df['Date'], df['Job'], df['notes']], axis=1)
    df = df.rename(columns={0:'Break_Start', 1:'Break_End'})
    df = df.loc[df['Day'] == day]
    df = df.loc[df['Area_Time'] == area_time]
    if area_time != 'W':
        df = df.drop_duplicates('Name')
    
    #Deletes Green Clean Entries 
    df['Name'].replace('', np.nan, inplace=True)
    df['Name'].replace('Green Clean', np.nan, inplace=True)
    df['Name'].replace('Green Clean ', np.nan, inplace=True)
    #Drops columns without Entries
    df.dropna(subset=['Name'], inplace=True)

    return(df)


def extractDigits(lst):
    res = []
    for el in lst:
        sub = [el]
        res.append(sub)      
    return(res)

def minor(cell):
    output = ''
    if '*' in cell:
        output = 'M'
    return(output)

def gui_handler(day, area_time, date, shift_start = '8:00am', shift_end = '5:00pm'):

    df = api_handler(day, area_time)
    shift = {'Shift_Start':[], 'Shift_End':[]}
    
    
    for name in df['Name']:
        
        if '-' in name:
            x = str(name[name.find('(')+1:name.find(')')]).split('-')
            shift['Shift_Start'].append(x[0])
            shift['Shift_End'].append(x[1])
        
        else:
            shift['Shift_Start'].append(shift_start)
            shift['Shift_End'].append(shift_end)
            
    df['Shift_Start'] = shift['Shift_Start']
    df['Shift_End'] = shift['Shift_End']
    df['Name'] = df['Name'].str.replace(r"\(.*\)","")
    df = df.sort_values(by=['Name'])
    sups = df.loc[df['Job'] == 'Location Supervisor']['Name'].to_list()
    tls = df.loc[df['Job'] == 'Team Lead']['Name'].to_list()
    df['Name'].replace(tls, np.nan, inplace=True)
    df.dropna(subset=['Name'], inplace=True)
    df['Name'].replace(sups, np.nan, inplace=True)
    df.dropna(subset=['Name'], inplace=True)
    df['Minor'] = df['Name'].apply(minor)
    
    # Interacts with Open Xl sheet
    new_book = xw.sheets.active
    new_book.range("A4:Q37").clear_contents()
    # Waterpark Exception
    if area_time == 'W':
        df1 = df
        df = df.drop_duplicates('Name', keep = 'first')
        df1 = df1.drop_duplicates('Name', keep = 'last')
        #Posts Jared Notes
        new_book.range('O7').value = df['notes']
        new_book.range("O7:O37").clear_contents()
        new_book.range('N7').value = df1['Break_End']
        new_book.range('M7').value = df1['Break_Start']
    else:    
        #Posts Jared Notes
        new_book.range('O7').value = df['notes']
        new_book.range("O7:O37").clear_contents()
    
    
    new_book.range('F1').value = area_time
    new_book.range('K1').value = date
    new_book.range('J8:K32').clear_contents()
    #posts second break
    new_book.range('J7').value = df['Break_End']
    new_book.range('I7').value = df['Break_Start']
    
    new_book.range('D7').value = df['Shift_End']
    new_book.range('C7').value = df['Shift_Start']
    new_book.range('C7:C32').api.Font.Size = 5
    new_book.range('B7').value = df['Job']
    new_book.range('A7').value = df['Name']
    list_ = df['Minor'].to_list()
    new_book.range('A8').value = extractDigits(list_)
    new_book.range("I7:I37").clear_contents()
    new_book.range("B7:P7").clear_contents()
    new_book.range("M7:M32").clear_contents()
    
    
    if len(sups) > 1:
        new_book.range('B4').value = 'Supervisor: %s' % sups[0]
        new_book.range('B5').value = 'Supervisor: %s' % sups[1]
    elif len(sups) > 0:
        new_book.range('B4').value = 'Supervisor: %s' % sups[0]
        new_book.range('B5').value = 'Supervisor:'
    else:
        new_book.range('B4:B5').value = 'Supervisor:'
              
    if len(tls) >1:  
        new_book.range('B6').value = 'Team Lead: %s' % tls[0]
        new_book.range('B7').value = 'Team Lead: %s' % tls[1]
    elif len(tls) >0:
        new_book.range('B6').value = 'Team Lead: %s' % tls[0]
        new_book.range('B7').value = 'Team Lead:'
    else:
        new_book.range('B6:B7').value = 'Team Lead:'
        
import tkinter as tk

class SampleApp(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.calendar = Calendar(self)
        self.title("Break Sheets")
        self.calendar.pack()
        self.area1_am_button = tk.Button(self, text="A1 AM", command=self.store_entry_a1_am)
        self.area1_am_button.pack()
        self.area1_am_button.place(x=5, y=186) 
        self.area1_pm_button = tk.Button(self, text="A1 PM", command=self.store_entry_a1_pm)
        self.area1_pm_button.pack()
        self.area1_pm_button.place(x=5, y=211)
        self.area2_am_button = tk.Button(self, text="A2 AM", command=self.store_entry_a2_am)
        self.area2_am_button.pack()
        self.area2_am_button.place(x=65, y=186)
        self.area2_pm_button = tk.Button(self, text="A2 PM", command=self.store_entry_a2_pm)
        self.area2_pm_button.pack()
        self.area2_pm_button.place(x=139, y=186)
        self.area3_am_button = tk.Button(self, text="A3 AM", command=self.store_entry_a3_am)
        self.area3_am_button.pack()
        self.area3_am_button.place(x=200, y=186)
        self.area3_pm_button = tk.Button(self, text="A3 PM", command=self.store_entry_a3_pm)
        self.area3_pm_button.pack()
        self.area3_pm_button.place(x=200, y=211)
        self.area4_button = tk.Button(self, text="A4", command=self.a4)
        self.area4_button.pack()
        self.w_button = tk.Button(self, text="W", command=self.w)
        self.w_button.pack()
        self.w_button.place(x=182, y=186)
        self.button = tk.Button(self, text="Click to Fill Break Sheet", command=self.on_button)
        self.button.pack()
        self.your_df = pd.DataFrame()
        
    def store_entry_a1_am(self):
        self.your_df = pd.DataFrame(data={'Date': [self.calendar.get_date()]})
        self.your_df['Area and Time'] = self.area1_am_button.cget('text')
        self.shift = ['8:00am', '5:00pm']
        shift_label = Label(self, text = "A1 AM").pack(side = LEFT)
        string_var = tk.StringVar(self, '%s - %s' % (self.shift[0], self.shift[1]))
        self.shift_start = tk.Entry(self, textvariable=string_var).pack(side = TOP)
    
    def store_entry_a1_pm(self):
        self.your_df = pd.DataFrame(data={'Date': [self.calendar.get_date()]})
        self.your_df['Area and Time'] = self.area1_pm_button.cget('text')
        self.shift = ['4:00pm', '2:00am']
        shift_label = Label(self, text = "A1 PM").pack(side = LEFT)
        string_var = tk.StringVar(self, '%s - %s' % (self.shift[0], self.shift[1]))
        self.shift_start = tk.Entry(self, textvariable=string_var).pack()

    def store_entry_a2_am(self):
        self.your_df = pd.DataFrame(data={'Date': [self.calendar.get_date()]})
        self.your_df['Area and Time'] = self.area2_am_button.cget('text')
        self.shift = ['8:00am', '5:00pm']
        shift_label = Label(self, text = "A2 AM").pack(side = LEFT)
        string_var = tk.StringVar(self, '%s - %s' % (self.shift[0], self.shift[1]))
        self.shift_start = tk.Entry(self, textvariable=string_var).pack()
        
    def store_entry_a2_pm(self):
        self.your_df = pd.DataFrame(data={'Date': [self.calendar.get_date()]})
        self.your_df['Area and Time'] = self.area2_pm_button.cget('text')
        self.shift = ['4:00pm', '2:00am']
        shift_label = Label(self, text = "A2 PM").pack(side = LEFT)
        string_var = tk.StringVar(self, '%s - %s' % (self.shift[0], self.shift[1]))
        self.shift_start = tk.Entry(self, textvariable=string_var).pack()

    def store_entry_a3_am(self):
        self.your_df = pd.DataFrame(data={'Date': [self.calendar.get_date()]})
        self.your_df['Area and Time'] = self.area3_am_button.cget('text')
        self.shift = ['8:00am', '5:00pm']
        shift_label = Label(self, text = "A3 AM").pack(side = LEFT)
        string_var = tk.StringVar(self, '%s - %s' % (self.shift[0], self.shift[1]))
        self.shift_start = tk.Entry(self, textvariable=string_var).pack(side = TOP)
        
    def store_entry_a3_pm(self):
        self.your_df = pd.DataFrame(data={'Date': [self.calendar.get_date()]})
        self.your_df['Area and Time'] = self.area3_pm_button.cget('text')
        self.shift = ['4:00pm', '2:00am']
        shift_label = Label(self, text = "A3 PM").pack(side = LEFT)
        string_var = tk.StringVar(self, '%s - %s' % (self.shift[0], self.shift[1]))
        self.shift_start = tk.Entry(self, textvariable=string_var).pack()    
    
    def a4(self):
        self.your_df = pd.DataFrame(data={'Date': [self.calendar.get_date()]})
        self.your_df['Area and Time'] = self.area4_button.cget('text')
        self.shift = ['4:00pm', '2:00am']
        shift_label = Label(self, text = "A4").pack(side = LEFT)
        string_var = tk.StringVar(self, '%s - %s' % (self.shift[0], self.shift[1]))
        self.shift_start = tk.Entry(self, textvariable=string_var)
        self.shift_start.pack()
    
    def w(self):
        self.your_df = pd.DataFrame(data={'Date': [self.calendar.get_date()]})
        self.your_df['Area and Time'] = self.w_button.cget('text')
        self.shift = ['4:00pm', '2:00am']
        shift_label = Label(self, text = "W").pack(side = LEFT)
        string_var = tk.StringVar(self, '%s - %s' % (self.shift[0], self.shift[1]))
        self.shift_start = tk.Entry(self, textvariable=string_var)
        self.shift_start.pack()
    
    
    def on_button(self):
        self.your_df['Date'] = pd.to_datetime(self.your_df['Date'], format="%m/%d/%y")
        self.your_df['Day of week (int)'] = self.your_df['Date'].dt.weekday
        self.your_df['Day of week (str)'] = self.your_df['Date'].dt.day_name()
        print(self.calendar.get_date())

        self.destroy()
        gui_handler(self.your_df['Day of week (str)'][0], self.your_df['Area and Time'][0], self.calendar.get_date(), self.shift[0], self.shift[1])

app = SampleApp()
app.mainloop()        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        