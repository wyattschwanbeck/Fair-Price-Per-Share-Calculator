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
        
def check_path_and_download(ticker):
    newpath = "{}/".format(ticker)     
    if not os.path.exists(newpath):
        os.makedirs(newpath)
        print("Creating folder for {}".format(ticker))
    fin_dl = FinancialsDownloader()
    print("Downloading financials for {}.".format(ticker))
    fins = fin_dl.download(ticker)   
    


userQuit = False
while(userQuit!=True):
    ticker = input("Enter Ticker Symbol or Type Q() to Quit:")
    if ticker=="Q()":
        userQuit=True
    else:
        check_path_and_download(ticker)
        income_statement = "{}/incomeStatement.csv".format(ticker, ticker)
        cash_flow = "{}/cashFlow.csv".format(ticker, ticker)
        balance_sheet = "{}/balanceSheet.csv".format(ticker, ticker)
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


            
        
