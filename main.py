# -*- coding: utf-8 -*-
"""
Created on Fri Apr 20 10:13:19 2018
main caller for FPPS generation
@author: Wyatt
"""
from Calculate_Value_Drivers import Calculate_Value_Drivers
from FPPS_Excel_Injector import FPPS_Excel_Injector
from MS_Financial_Downloader import FinancialsDownloader
import os
        
ticker = "DIS"

from tkinter import *
import os  

class Window(Frame):
    
    def __init__(self, master = None):
        Frame.__init__(self, master)
        
        self.master = master
        self.tickerEntry = None
        self.init_window()
        
        
    def init_window(self):
        self.master.title("Fair Price Per Share Generator")
        #adjust dimensions, expand fills
        self.pack(fill = BOTH, expand=1)
        
        
        self.tickerEntry = Entry(self.master)
        self.tickerEntry.place(x=80, y=25)
        self.tickerEntry.delete(0, END)
        self.tickerEntry.insert(0,"DIS")
        
        
        RunButton = Button\
        (self, 
         text="Produce Modeled Fair Price Per Share",
         command=self.run_FPPS)
        RunButton.place(x=40, y=75)

        
        menu = Menu(self.master)
        self.master.config(menu = menu)
        
        file = Menu(menu)
        file.add_command(label = "Exit", command = self.client_exit)
        menu.add_cascade(label = "File", menu = file)
        
        edit = Menu(menu)
        edit.add_command(label = "Reset Defaults")
        menu.add_cascade(label = "Edit", menu = edit)


    def run_FPPS(self):
        ticker = self.tickerEntry.get()
        check_path_and_download(ticker)
        perform_calculations_and_excel_injection(ticker)
        
    def client_exit(self):
        self.master.destroy()
        
def check_path_and_download(ticker):
    newpath = "{}/".format(ticker)     
    if not os.path.exists(newpath):
        os.makedirs(newpath)
        print("Creating folder for {}".format(ticker))
    fin_dl = FinancialsDownloader()
    print("Downloading financials for {}.".format(ticker))
    fins = fin_dl.download(ticker)   
    fins["income_statement"].to_csv\
    ("{}/{} Income Statement.csv".format(ticker,ticker),index=False,mode="w")
    fins["cash_flow"].to_csv\
    ("{}/{} Cash Flow.csv".format(ticker,ticker), index = False, mode = "w")
    fins["balance_sheet"].to_csv\
    ("{}/{} Balance Sheet.csv".format(ticker,ticker), index = False, mode = "w")
    

def perform_calculations_and_excel_injection(ticker):
    income_statement = "{}/{} Income Statement.csv".format(ticker, ticker)
    cash_flow = "{}/{} Cash Flow.csv".format(ticker, ticker)
    balance_sheet = "{}/{} Balance Sheet.csv".format(ticker, ticker)
    value_drivers = Calculate_Value_Drivers\
    (income_statement, cash_flow, balance_sheet, ticker, .015)
    value_drivers.calculate_value_drivers()
    value_drivers.write_to_csv()
    
    #b = FPPS.Free_Cash_Flow(a, 2, ticker)
    c = FPPS_Excel_Injector\
    (value_drivers, 
    "{}/{} Compiled Projection.xlsx".format(ticker,ticker), 
    "{}/".format(ticker), 
    years_to_project= 5)
    
    c.load_financial_data_to_sheets()
    
    
    c.workbook.close()
        
root = Tk()
root.geometry("300x150")


app = Window(root)

root.mainloop()

