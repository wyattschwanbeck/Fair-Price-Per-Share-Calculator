"""
Created on Fri Mar  9 09:22:09 2018
Using Past 5 year, income statement, statement of cash flows, and balance sheet downloads from Morningstar.com
@author: Wyatt
"""
import pandas as pd
import numpy as np
import get_stock_data
from market_returns import market_data

class Calculate_Value_Drivers(object):
    def __init__(self, income_statement, cash_flows, balance_sheet, ticker, risk_free_rate, ticker_returns):
        #import CSV data with panda, and replace "NaN" datas
        self.value_drivers_list = ["Sales Growth", \
        "Operating Expenses to Sales", "Depreciation", \
        "Operating Assets to Sales", \
        "Operating Current Liabilities to Sales", \
        "Capital Expenditure to Sales", \
        "Interest earned on cash and equivalents", 
        "Cost of debt", \
        "Tax Rate", \
        "ONWC", \
        "Average long-term debt", \
        "Shares outstanding", \
        "Cost of equity", \
        "Risk free rate", \
        "Beta",
        "WACC"]        
        self.value_drivers_dict = {"Sales Growth" :0, \
        "Operating Expenses to Sales" : 0, \
        "Depreciation" : 0, \
        "Operating Current Assets to Sales": 0, \
        "Operating Current Liabilities to Sales": 0, \
        "Capital Expenditure to Sales" : 0, \
        "Interest earned on cash and equivalents" : risk_free_rate, 
        "Cost of debt" : 0, \
        "Tax Rate" : 0,
        "ONWC" : 0,
        "Average long-term debt" : 0,
        "Shares outstanding" : 0,
        "Cost of equity" : 0,
        "Risk free rate" : risk_free_rate,
        "Beta" : 0,\
        "WACC" : 0}

        self.ticker_return = ticker_returns.average_yearly_return
        
        
        
        self.sheet_index = {"IS" : dict(), "CF" :dict(), "BS":dict()}
        self.IS = pd.read_csv(income_statement, skiprows=1)
        self.IS.replace("NaN", 0, inplace = True)
        self.CF = pd.read_csv(cash_flows, skiprows=1)
        self.CF.replace("NaN", 0, inplace = True)
        self.BS = pd.read_csv(balance_sheet, skiprows=1)
        self.BS.replace("NaN", 0, inplace = True)

        self.IS = np.array(self.IS)
        #self.verify_names(self.IS, "Financial Statement Data Labels/Income Statement.txt")
        self.BS = np.array(self.BS)
        self.CF = np.array(self.CF)
        
        self.generate_index_lists()
        
        self.calculate_sales_growth()
        self.calculate_operating_exp_to_sales()
        self.calculate_depreciation()
        self.calculate_operating_assets_to_sales()
        self.calculate_operating_liabilities_to_sales()
        self.capital_expenditure_to_sales()
        self.interest_paid_on_debt()
        self.calculate_tax_rate()
        self.calculate_ONWC()
        self.calculate_average_debt()
        
        self.raw_stock_data = get_stock_data.parse(ticker)
        self.calculate_shares_outstanding()
        self.calculate_cost_of_equity()
        self.calculate_WACC()
    
    def verify_data_labels(self, sheet_array, data_labels_txt):
        names = pd.read_csv("data_labels_txt")
        names = np.array(names)
        
    def generate_index_lists(self):
        for r,rows in enumerate(self.IS,0):
            self.sheet_index["IS"].update({rows[0] : r})
        for r,rows in enumerate(self.CF,0):
            self.sheet_index["CF"].update({rows[0] : r})        
        for r,rows in enumerate(self.BS,0):
            self.sheet_index["BS"].update({rows[0] : r})
  
    def average_rate_of_change(self, row):
        rates = []
        for y,year in enumerate(row[1:-1], 0):
            if(y == 0):
                previous_year = year  
                continue
            else:
                latest_year = year                
                rates.append((latest_year - previous_year)/previous_year)
                previous_year = latest_year
        return np.average(rates)
        
    def calculate_sales_growth(self):
        for row in self.IS:
            if(row[0] == "Revenue"):
                sales_growth = self.average_rate_of_change(row)
                
                self.value_drivers_dict["Sales Growth"] = round(sales_growth, 4)
                print("Sales Growth: {}".format(self.value_drivers_dict["Sales Growth"]))
                
    def calculate_operating_exp_to_sales(self):
        '''excluding depreciation from operating expenses,
        average((cost of revenue + total operating expenses)/sales)'''
        yearly = []
        for year in range(1,6):
            yearly.append(\
            (self.IS[self.sheet_index["IS"]["Cost of revenue"]][year] + \
            self.IS[self.sheet_index["IS"]["Total costs and expenses"]][year]) / \
            self.IS[self.sheet_index["IS"]["Revenue"]][year])
        print("Operating Expenses to Sales:{}".format(np.average(yearly)))
        self.value_drivers_dict["Operating Expenses to Sales"] = round(np.average(yearly), 4)
    def calculate_depreciation(self):
        yearly = []
        for year in range(1,6):
            yearly.append(\
            (self.IS[self.sheet_index["CF"]["Depreciation & amortization"]][year] / \
            self.BS[self.sheet_index["BS"]["Gross property, plant and equipment"]][year]))
        print("Depreciation: {}".format(np.average(yearly)))
        self.value_drivers_dict["Depreciation"] = round(np.average(yearly), 4)
        
    def calculate_operating_assets_to_sales(self):
        yearly = []
        for year in range(1,6):
            yearly.append(\
            (self.BS[self.sheet_index["BS"]["Receivables"]][year] + \
            self.BS[self.sheet_index["BS"]["Inventories"]][year])/ \
            self.IS[self.sheet_index["IS"]["Revenue"]][year])
        print("Operating Assets to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["Operating Current Assets to Sales"] = round(np.average(yearly), 4)
    
    def calculate_operating_liabilities_to_sales(self):
        yearly = []
        for year in range(1,6):
            yearly.append(\
            (self.BS[self.sheet_index["BS"]["Total current liabilities"]][year])/ \
            self.IS[self.sheet_index["IS"]["Revenue"]][year])
        print("Operating Liabilities to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["Operating Current Liabilities to Sales"] = round(np.average(yearly), 4)    
        
    def capital_expenditure_to_sales(self):
        yearly = []
        for year in range(1,6):
            yearly.append(\
            (self.BS[self.sheet_index["CF"]["Capital expenditure"]][year])/ \
            self.IS[self.sheet_index["IS"]["Revenue"]][year])
        print("Capital Expenditure to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["Capital Expenditure to Sales"] = round(np.average(yearly), 4)
        
        
    def interest_paid_on_debt(self):
        rates = []
        for y,year in enumerate(self.BS[self.sheet_index["BS"]["Long-term debt"]][1:6], 0):
            if(y == 0):
                previous_year = year  
                continue
            else:
                latest_year = year                
                rates.append\
                (self.IS[self.sheet_index["IS"]["Interest expense"]][y+1]/((latest_year + previous_year)/2))
                previous_year = latest_year
        print("Interest paid on Debt: {}".format(np.average(rates)))
        self.value_drivers_dict["Cost of debt"] = np.average(rates)
        return np.average(rates)

    def calculate_tax_rate(self):
        yearly = []
        for year in range(1,6):
            yearly.append(\
            (self.IS[self.sheet_index["IS"]["Provision for income taxes"]][year])/ \
            self.IS[self.sheet_index["IS"]["Income before income taxes"]][year])
        print("Tax Rate: {}".format(np.average(yearly)))
        self.value_drivers_dict["Tax Rate"] = round(np.average(yearly), 4)
        
    def calculate_ONWC(self):
        yearly = []
        for year in range(1,6):
            yearly.append(\
            ((self.BS[self.sheet_index["BS"]["Receivables"]][year] + \
            self.BS[self.sheet_index["BS"]["Inventories"]][year]) - \
            self.BS[self.sheet_index["BS"]["Total current liabilities"]][year])/ \
            self.IS[self.sheet_index["IS"]["Revenue"]][year])
        print(np.average(yearly))
        self.value_drivers_dict["ONWC"] = round(np.average(yearly), 4)    
        
    def calculate_average_debt(self):
        yearly = []
        for year in range(1,6):
            yearly.append(\
            self.BS[self.sheet_index["BS"]["Long-term debt"]][year])
        print("Average Long Term Debt: {}".format(np.average(yearly)))
        self.value_drivers_dict["Average long-term debt"] = round(np.average(yearly), 4)
        
    def calculate_shares_outstanding(self):
        '''Market Cap / share price '''
        market_cap = self.raw_stock_data["Market Cap"]    
        if(market_cap[-1] == "B"):
            market_cap = int(float(market_cap[0:-1])*1000000000)
        elif(market_cap[-1] == "M"):
            market_cap = int(float(market_cap[0:-1])*1000000)
        price = float(self.raw_stock_data["Open"])
        self.value_drivers_dict["Shares outstanding"] = int(market_cap/price)
        print("Shares Outstanding : {}".format(self.value_drivers_dict["Shares outstanding"]))
    
    def calculate_cost_of_equity(self):
        '''CAPM baby. Cost of Equity = risk_free + (Beta*(Market_returns - risk_free))'''
        if(self.raw_stock_data["Beta"] != "N/A"):            
            self.value_drivers_dict["Beta"] = float(self.raw_stock_data["Beta"])
            self.value_drivers_dict["Cost of equity"] = self.value_drivers_dict["Risk free rate"] + (self.value_drivers_dict["Beta"] *(self.value_drivers_dict["Risk free rate"]))
            print("Cost of Equity: {}".format(self.value_drivers_dict["Cost of equity"]))
        else:
            print("Do Gordon Growth Formula")
                    
            
    def calculate_WACC(self):
        '''(Cost of debt *((Debt / (Equity + Debt)) * (1-Tax rate))) + (Cost of equity)*((equity)/(equity + Debt))'''
        current_price = float(self.raw_stock_data["Open"])
        cost_debt = self.value_drivers_dict["Cost of debt"]
        cost_equity = self.value_drivers_dict["Cost of equity"]
        avg_debt = self.value_drivers_dict["Average long-term debt"]
        tax_rate = self.value_drivers_dict["Tax Rate"]
        #convert shares outstanding to millions to match financial statement data format
        total_equity = (self.value_drivers_dict["Shares outstanding"]/1000000) *\
            current_price
            
        self.value_drivers_dict["WACC"] = (cost_debt *((avg_debt / (total_equity + avg_debt)) * (1-tax_rate))) + (cost_equity)*((total_equity)/(total_equity + avg_debt))
        print("WACC : {}".format(self.value_drivers_dict["WACC"]))
   
