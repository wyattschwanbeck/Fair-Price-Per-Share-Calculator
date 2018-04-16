"""
Created on Fri Mar  9 09:22:09 2018
Updated 4/16/2018, 10:10 AM EST
Using Past 5 year, income statement, statement of cash flows, and balance sheet downloads from Morningstar.com
@author: Wyatt
"""
import pandas as pd
import numpy as np
import get_stock_data
from FPPS_Excel_Injector import FPPS_Excel_Injector
import csv

class Calculate_Value_Drivers(object):
    def __init__(self, \
    income_statement, \
    cash_flows, \
    balance_sheet, \
    ticker, \
    risk_free_rate):
        #import CSV data with panda, and replace "NaN" datas
        self.ticker = ticker        
        self.value_drivers_list = ["Sales Growth", \
        "Operating Expenses to Sales", "Depreciation", \
        "Operating Current Assets to Sales", \
        "Operating Current Liabilities to Sales", \
        "Capital Expenditure to Sales", \
        "Interest earned on cash and equivalents", 
        "Cost of debt", \
        "Tax Rate", \
        "ONWC", \
        "Latest Long-term debt", \
        "Shares outstanding", \
        "Cost of equity", \
        "Risk free rate", \
        "Beta",
        "WACC",
        "Long Term Growth",
        "Market Return",
        "Current Share Price",
        "Latest ONWC",
        "Dividend Rate",
        "Dividend Growth Rate",
        "Annual Stock Repurchase"]        
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
        "Latest Long-term debt" : 0,
        "Shares outstanding" : 0,
        "Cost of equity" : 0,
        "Risk free rate" : risk_free_rate,
        "Beta" : 0,\
        "WACC" : 0,
        "Long Term Growth" : 0.0125,
        "Market Return" : .12,
        "Current Share Price" : 0,\
        "Dividend Rate" : 0,
        "Latest ONWC" : 0,
        "Dividend Rate" : 0,
        "Dividend Growth Rate" : 0,
        "Annual Stock Repurchase" : 0}

        
        
        
        self.sheet_index = {"IS" : dict(), "CF" :dict(), "BS":dict()}
        self.IS = pd.read_csv(income_statement, skiprows=1)
        self.IS.replace("NaN", 0, inplace = True)
        

        self.CF = pd.read_csv(cash_flows, skiprows=1)
        self.CF.replace("NaN", 0, inplace = True)
        self.BS = pd.read_csv(balance_sheet, skiprows=1)
        self.BS.replace("NaN", 0, inplace = True)
        
        self.latest_year = self.BS.columns.values[-1][0:4]
        self.years = self.BS.columns.values[1:]    

        self.IS = np.array(self.IS)
        #self.verify_names(self.IS, "Financial Statement Data Labels/Income Statement.txt")
        self.BS = np.array(self.BS)
        self.CF = np.array(self.CF)
        self.IS = \
            self.verify_data_labels\
            (self.IS, "Financial Statement Data Labels/Income Statement.csv")
        self.BS = \
            self.verify_data_labels\
            (self.BS, "Financial Statement Data Labels/Balance Sheet.csv")
        

    def calculate_value_drivers(self):
        self.generate_index_lists()
        
        self.calculate_sales_growth()
        self.calculate_operating_exp_to_sales()
        self.value_drivers_dict["Depreciation"] = .40
        #self.calculate_depreciation()        
        self.calculate_operating_assets_to_sales()
        self.calculate_operating_liabilities_to_sales()
        self.capital_expenditure_to_sales()
        self.calculate_latest_debt()
        self.interest_paid_on_debt()
        #self.calculate_tax_rate()
        self.value_drivers_dict["Tax Rate"] = .21      
        self.last_ONWC = 0     
        self.calculate_ONWC()        
        
        self.raw_stock_data = get_stock_data.parse(ticker)
        self.calculate_shares_outstanding()
        self.calculate_cost_of_equity()
        self.calculate_WACC()
        self.calculate_dividend_rate()
        self.calculate_average_stock_repurchase()
        
    
    def verify_data_labels(self, sheet_array, data_labels_csv):
        names = pd.read_csv(data_labels_csv)
        names = np.array(names)
        for n,name in enumerate(names):
            
            for c,check_name in enumerate(sheet_array[0:]):
                            
                if(check_name[0] in name and name[0]!=check_name[0]):
                    print("Replacing Data Label: {} to {}".format(check_name[0], name[0]))
                    check_name[0] = name[0]
        
        return sheet_array
        
    def get_BS_item(self, item, index):
        try:
            return self.BS[self.sheet_index["BS"][item]][index]
        except KeyError:
            print("Unable to retreive {} from BS at index {}".format(item, index))
            return 0

    def get_IS_item(self, item, index):
        try:
            if(index == False):
                return self.IS[self.sheet_index["IS"][item]]
                
            return self.IS[self.sheet_index["IS"][item]][index]
        except KeyError:
            print("Unable to retreive {} from IS at index {}".format(item, index))
            return 0
    
    def get_CF_item(self, item, index, make_positive = False):
        try:
            item = self.CF[self.sheet_index["CF"][item]][index]
            if(make_positive == True and item < 0):    
                return -item
            else:
                return item
                
        except KeyError:
            print("Unable to retreive {} from IS at index {}".format(item, index))
            return 0
        
    def generate_index_lists(self):
        for r,rows in enumerate(self.IS,0):
            if(rows[0] == "Total net revenue"):
                self.sheet_index["IS"]["Revenue"] = r
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
        rev = self.IS[self.sheet_index["IS"]["Revenue"]]
        sales_growth = self.average_rate_of_change(rev)
                
        self.value_drivers_dict["Sales Growth"] = round(sales_growth, 4)
        print("Sales Growth: {}".format(self.value_drivers_dict["Sales Growth"]))
                
    def calculate_operating_exp_to_sales(self):
        '''excluding depreciation from operating expenses,
        average((cost of revenue + total operating expenses)/sales)'''
        yearly = []
        for year in range(1,6):
            cor = self.get_IS_item("Cost of revenue", year)            
            op_costs = self.get_IS_item("Total costs and expenses", year)
            rev = self.get_IS_item("Revenue", year)
            yearly.append(\
            (cor + \
            op_costs) / \
            rev)
        print("Operating Expenses to Sales:{}".format(np.average(yearly)))
        self.value_drivers_dict["Operating Expenses to Sales"] =\
        round(np.average(yearly), 4)
    
    '''def calculate_depreciation(self):
        yearly = []
        depreciation = self.CF[self.sheet_index["CF"]["Depreciation & amortization"]]
        ppe = self.BS[self.sheet_index["BS"]['Gross property, plant and equipment']]
        for year in range(4,6):
            yearly.append(\
            (depreciation[year]/ \
            ppe[year]))
        print("Depreciation: {}".format(np.average(yearly)))
        self.value_drivers_dict["Depreciation"] = round(np.average(yearly), 4)'''
        
        
    def calculate_operating_assets_to_sales(self):
        yearly = []

        for year in range(1,6):
            rec = self.get_BS_item("Receivables", year)
            inv = self.get_BS_item("Inventories", year)
            rev = self.get_IS_item("Revenue", year)
            value_driver = (rec + inv)

            yearly.append(value_driver/rev)
        print("Operating Assets to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["Operating Current Assets to Sales"] = \
            round(np.average(yearly), 4)
    
    def calculate_operating_liabilities_to_sales(self):
        yearly = []
        for year in range(1,6):
            tc_liab = self.get_BS_item("Total current liabilities", year)
            rev = self.get_IS_item("Revenue", year)
            yearly.append(tc_liab/rev)
        print("Operating Liabilities to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["Operating Current Liabilities to Sales"] =\
        round(np.average(yearly), 4)    
        
    def capital_expenditure_to_sales(self):
        yearly = []
        for year in range(1,6):
            ce = self.get_CF_item("Capital expenditure", year, make_positive = True)
            rev = self.get_IS_item("Revenue", year)            
            yearly.append(ce/rev)
        print("Capital Expenditure to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["Capital Expenditure to Sales"] = \
        round(np.average(yearly), 4)
        
        
            
            
    def interest_paid_on_debt(self):
        rates = []
        try:
            ltd = self.BS[self.sheet_index["BS"]["Long-term debt"]][1:6]
            for y,year in enumerate(ltd, 0):
                if(y == 0):
                    previous_year = year  
                    continue
                else:
                    latest_year = year                
                    try:
                        rates.append\
                    (self.IS[self.sheet_index["IS"]["Interest expense"]][y+1]/\
                    ((latest_year + previous_year)/2))
                    except ZeroDivisionError:
                        rates.append(0)
                    except KeyError:
                        rates.append(0)
                    previous_year = latest_year
            print("Interest paid on Debt: {}".format(np.average(rates)))
            self.value_drivers_dict["Cost of debt"] = np.average(rates)
            return np.average(rates)
        except KeyError:
            print("Long Term Debt Not Found!")
            self.value_drivers_dict["Cost of debt"] = 0

    def calculate_tax_rate(self):
        yearly = []
        for year in range(1,6):
            if(self.IS[self.sheet_index["IS"]["Provision for income taxes"]][year]<0):
                yearly.append(\
                (-self.IS[self.sheet_index["IS"]["Provision for income taxes"]][year])/ \
                self.IS[self.sheet_index["IS"]["Income before income taxes"]][year])
            else:
                yearly.append(\
                (self.IS[self.sheet_index["IS"]["Provision for income taxes"]][year])/ \
                self.IS[self.sheet_index["IS"]["Income before income taxes"]][year])
        print("Tax Rate: {}".format(np.average(yearly)))
        self.value_drivers_dict["Tax Rate"] = round(np.average(yearly), 4)
        
    def calculate_ONWC(self):
        yearly = []            
        for year in range(1,6):
            rec = self.get_BS_item("Receivables", year)
            inv = self.get_BS_item("Inventories", year)
            yearly.append(\
            ((rec + inv) - \
            self.BS[self.sheet_index["BS"]["Total current liabilities"]][year])/ \
            self.IS[self.sheet_index["IS"]["Revenue"]][year])
            if(year == 5):
               self.value_drivers_dict["Latest ONWC"] = (rec + \
            inv) - \
            self.BS[self.sheet_index["BS"]["Total current liabilities"]][year]
        
        print("Operating Net Working Capital to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["ONWC"] = round(np.average(yearly), 4)    
        
    def calculate_latest_debt(self):
        try:
            lt = self.BS[self.sheet_index["BS"]["Long-term debt"]][-1]
            print("Latest Long-term Debt: {}".format(lt))
            self.value_drivers_dict["Latest Long-term debt"] = round(lt, 4)
        except KeyError:
            print("Latest Long term Debt: 0")
        
    def calculate_shares_outstanding(self):
        '''Market Cap / share price '''
        try:        
            market_cap = self.raw_stock_data["Market Cap"]  
        except:
            market_cap = "233.903B"
        if(market_cap[-1] == "B"):
            market_cap = int(float(market_cap[0:-1])*1000000000)
        elif(market_cap[-1] == "M"):
            market_cap = int(float(market_cap[0:-1])*1000000)
        try:
            price = float(self.raw_stock_data["Open"].replace(",",""))
        except:
            price = 76.5
        self.value_drivers_dict["Shares outstanding"] = int(market_cap/price)
        print("Shares Outstanding : {}".format(self.value_drivers_dict["Shares outstanding"]))
    
    def calculate_cost_of_equity(self):
        '''CAPM baby. Cost of Equity = risk_free + (Beta*(Market_returns - risk_free))'''
        try:        
            b = self.raw_stock_data["Beta"] != "N/A"
        except:
            b = "N/A"
        if(b != "N/A"):            
            self.value_drivers_dict["Beta"] = float(self.raw_stock_data["Beta"])
            self.value_drivers_dict["Cost of equity"] = self.value_drivers_dict["Risk free rate"] + (self.value_drivers_dict["Beta"] *(.12 - self.value_drivers_dict["Risk free rate"]))
            print("Cost of Equity: {}".format(self.value_drivers_dict["Cost of equity"]))
        else:
            print("Do Gordon Growth Formula")
                    
            
    def calculate_WACC(self):
        '''(Cost of debt *((Debt / (Equity + Debt)) * (1-Tax rate))) + (Cost of equity)*((equity)/(equity + Debt))'''
        current_price = float(self.raw_stock_data["Open"].replace(",", ""))
        self.value_drivers_dict["Current Share Price"] = current_price        
        cost_debt = self.value_drivers_dict["Cost of debt"]
        cost_equity = self.value_drivers_dict["Cost of equity"]
        lt_debt = self.get_BS_item("Long-term debt", -1)
        tax_rate = self.value_drivers_dict["Tax Rate"]
        #convert shares outstanding to millions to match financial statement data format
        total_equity = (self.value_drivers_dict["Shares outstanding"]/1000000) *\
            current_price
            
        self.value_drivers_dict["WACC"] = \
        (cost_debt *((lt_debt / (total_equity + lt_debt)) \
        * (1-tax_rate))) + (cost_equity)\
        *((total_equity)/(total_equity + lt_debt))
        print("WACC : {}".format(self.value_drivers_dict["WACC"]))
        
    def calculate_dividend_rate(self):
        dividend_rate = []
        dividend_growth = []
        last_div = 0
        for y,year in enumerate(self.years,1):
            div = self.get_CF_item("Dividend paid", y, make_positive = True)
            ni = self.get_IS_item("Net income", y)
            
            dividend_rate.append(div/ni)
       
            if(y == 1):
                dividend_growth.append(0)
            elif(last_div > 0):
                growth= (div - last_div)/last_div
                dividend_growth.append(growth)
            else:
                dividend_growth.append(0)
            
            last_div = div   
        print("Dividend Rate: {}".format(np.average(dividend_rate)))
        print("Dividend Growth: {}".format(np.average(dividend_growth)))
        self.value_drivers_dict["Dividend Rate"] = np.average(dividend_rate)
        self.value_drivers_dict["Dividend Growth Rate"] = np.average(dividend_growth)
        #for year in range(1,6):
    
    def calculate_average_stock_repurchase(self):
        stock_repurchase = []
        for y,year in enumerate(self.years,1):
            spent = \
            self.get_CF_item("Common stock repurchased",y,make_positive = True)
            stock_repurchase.append(spent)
        self.value_drivers_dict["Annual Stock Repurchase"] = \
            np.average(stock_repurchase)
            
        
        
    def value_driver(self, value_driver_name):
        try:
            return(self.value_drivers_dict[value_driver_name])
        except KeyError:
            print("Sorry that isn't in the value drivers list")
            
    def write_to_csv(self):
        with open("{}/{} Value Drivers.csv".format(self.ticker, self.ticker), "w") as f:
            writer = csv.writer(f, lineterminator='\n')
            
            for driver in self.value_drivers_list:
                writer.writerow([driver, self.value_drivers_dict[driver]])
    


        
        
        
        
ticker = "SBUX"
income_statement = "{}/{} Income Statement.csv".format(ticker, ticker)
cash_flow = "{}/{} Cash Flow.csv".format(ticker, ticker)
balance_sheet = "{}/{} Balance Sheet.csv".format(ticker, ticker)
value_drivers = Calculate_Value_Drivers(income_statement, cash_flow, balance_sheet, ticker, .015)
value_drivers.calculate_value_drivers()
value_drivers.write_to_csv()

#b = FPPS.Free_Cash_Flow(a, 2, ticker)
c = FPPS_Excel_Injector(value_drivers, "{}/{} Compiled Projection.xlsx".format(ticker,ticker), "{}/".format(ticker), 2)

c.load_financial_data_to_sheets()


c.workbook.close()



