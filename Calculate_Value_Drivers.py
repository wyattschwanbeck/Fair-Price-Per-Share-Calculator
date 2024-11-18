import pandas as pd
import numpy as np
from get_stock_data import get_stock_data
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
        "Latest Long-term debt",
        "Tax Rate", \
        "ONWC to Sales", \
        "Latest ONWC",
        "Cost of equity", \
        "Risk free rate", \
        "Beta",
        "WACC",
        "Long Term Growth",
        "Market Return",
        "Current Share Price",
        "Shares outstanding",
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
        "ONWC to Sales" : 0,
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
        self.IS = pd.read_csv(income_statement, skiprows=0)
        self.IS.replace("NaN", 0, inplace = True)


        self.CF = pd.read_csv(cash_flows, skiprows=0)
        self.CF.replace("NaN", 0, inplace = True)
        self.BS = pd.read_csv(balance_sheet, skiprows=0)
        self.BS.replace("NaN", 0, inplace = True)

        self.latest_year = self.BS.columns.values[-1][0:4]
        self.years = self.BS.columns.values[6:]

        self.IS = np.array(self.IS)
        self.BS = np.array(self.BS)
        self.CF = np.array(self.CF)
        self.IS = \
            self.verify_data_labels\
            (self.IS, "Financial Statement Data Labels/Income Statement.csv")
        self.BS = \
            self.verify_data_labels\
            (self.BS, "Financial Statement Data Labels/Balance Sheet.csv")


    def calculate_value_drivers(self):
        self._generate_index_lists()

        self.calculate_sales_growth()
        self.calculate_operating_exp_to_sales()
        self.value_drivers_dict["Depreciation"] = .30
        #self.calculate_depreciation()
        self.calculate_operating_assets_to_sales()
        self.calculate_operating_liabilities_to_sales()
        self.capital_expenditure_to_sales()
        self.calculate_latest_debt()
        self.interest_paid_on_debt()
        self.calculate_tax_rate()
        #self.value_drivers_dict["Tax Rate"] = .15
        self.last_ONWC = 0
        self.calculate_ONWC()

        self.raw_stock_data = get_stock_data(self.ticker)
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
                    print("Replacing Data Label: {} to {}".\
                    format(check_name[0], name[0]))
                    check_name[0] = name[0]

        return sheet_array

    def get_BS_item(self, item, index=False, make_positive = False):
        try:
            if(index != False):
                item = self.BS[self.sheet_index["BS"][item]][index]
            else:
                item = self.BS[self.sheet_index["BS"][item]]
            if(make_positive == True and item < 0 and index != False):
                return -item
            else:
                return item

        except KeyError:
            if(index == False):
                print("Unable to retreive list of {} from BS".\
            format(item))
                return [0]*len(self.BS[1])

            if(item == "Total Current Liabilities"):
                current_liab = ["Short-term borrowing",
                "Payables and accrued expenses",
                "Taxes payable"]
                print("Total current liabilities not given, calculating...")
                total_cur_liab = 0
                for liab in current_liab:
                    total_cur_liab += self.get_BS_item(liab, index)
                return total_cur_liab
            print("Unable to retreive {} from BS at index {}".\
            format(item, index))
            return 0

    def get_IS_item(self, item, index=False, make_positive = False):
        try:
            if(index != False):
                item = self.IS[self.sheet_index["IS"][item]][index]
            else:
                item = self.IS[self.sheet_index["IS"][item]]
            if(make_positive == True and item < 0 and index != False):
                return -item
            else:
                return item

        except KeyError:
            print("Unable to retreive {} from IS at index {}".\
            format(item, index))
            return 0

    def get_CF_item(self, item, index=False, make_positive = False):
        try:
            if(index != False):
                item = self.CF[self.sheet_index["CF"][item]][index]
            else:
                item = self.CF[self.sheet_index["CF"][item]]
            if(make_positive == True and item < 0 and index != False):
                return -item
            else:
                return item

        except KeyError:
            print("Unable to retreive {} from CF at index {}".\
            format(item, index))
            return 0

    def _generate_index_lists(self):
        for r,rows in enumerate(self.IS,0):
            if(rows[0] == "Total net revenue"):
                self.sheet_index["IS"]["Business Revenue"] = r
            self.sheet_index["IS"].update({rows[0] : r})
        for r,rows in enumerate(self.CF,0):
            self.sheet_index["CF"].update({rows[0] : r})
        for r,rows in enumerate(self.BS,0):
            self.sheet_index["BS"].update({rows[0] : r})

    def average_rate_of_change(self, row):
        rates = []
        for y,year in enumerate(row[6:-1], 0):
            if(y == 0):
                previous_year = year
                continue
            else:
                latest_year = year
                rates.append((latest_year - previous_year)/previous_year)
                previous_year = latest_year
        return np.average(rates)

    def calculate_sales_growth(self):
        rev = self.IS[self.sheet_index["IS"]["Business Revenue"]]
        sales_growth = self.average_rate_of_change(rev)

        self.value_drivers_dict["Sales Growth"] = round(sales_growth, 4)
        print("Sales Growth:{}".format(self.value_drivers_dict["Sales Growth"]))

    def calculate_operating_exp_to_sales(self):
        '''excluding depreciation from operating expenses,
        average((cost of revenue + total operating expenses)/sales)'''
        yearly = []
        for year in range(6,11):
            cor = self.get_IS_item("Cost of Revenue", year)
            #op_costs = self.get_IS_item("Total costs and expenses", year)
            rev = self.get_IS_item("Business Revenue", year)
            yearly.append((cor) / rev)
        print("Operating Expenses to Sales:{}".format(np.average(yearly)))
        self.value_drivers_dict["Operating Expenses to Sales"] =\
        round(np.average(yearly), 4)

    '''def calculate_depreciation(self):
        yearly = []
        depreciation = self.get_CF_item\
        ("Depreciation & amortization", index = False)
        ppe = self.get_BS_item\
        ('Gross property, plant and equipment', index = False)
        for year in range(4,6):
            yearly.append(\
            (depreciation[year]/ \
            ppe[year]))
        print("Depreciation: {}".format(np.average(yearly)))
        self.value_drivers_dict["Depreciation"] = round(np.average(yearly), 4)'''


    def calculate_operating_assets_to_sales(self):
        yearly = []

        for year in range(6,11):
            rec = self.get_BS_item("Trade and Other Receivables, Current", year)
            inv = self.get_BS_item("Inventories", year)
            rev = self.get_IS_item("Business Revenue", year)
            value_driver = (rec + inv)

            yearly.append(value_driver/rev)
        print("Operating Assets to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["Operating Current Assets to Sales"] = \
            round(np.average(yearly), 4)

    def calculate_operating_liabilities_to_sales(self):
        yearly = []
        for year in range(6,11):
            tc_liab = self.get_BS_item("Total Current Liabilities", year)
            rev = self.get_IS_item("Business Revenue", year)
            yearly.append(tc_liab/rev)
        print("Operating Liabilities to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["Operating Current Liabilities to Sales"] =\
        round(np.average(yearly), 4)

    def capital_expenditure_to_sales(self):
        yearly = []
        for year in range(6,11):
            
            ce = self.get_CF_item\
            ("Capital Expenditure, Reported", year, make_positive = True)
            if ce == 0:
                ce = self.get_CF_item\
                ("Purchase/Sale and Disposal of Property, Plant and Equipment, Net", year, make_positive = True)
            rev = self.get_IS_item("Business Revenue", year)
            yearly.append(ce/rev)
        print("Capital Expenditure to Sales: {}".format(np.average(yearly)))
        self.value_drivers_dict["Capital Expenditure to Sales"] = \
        round(np.average(yearly), 4)




    def interest_paid_on_debt(self):
        rates = []
        try:
            ltd = self.get_BS_item("Long Term Debt")[6:11]
            for y,year in enumerate(ltd, 0):
                int_ex = self.get_IS_item("Interest Expense Net of Capitalized Interest",y+5+1)
                if(y == 0):
                    previous_year = year
                    continue
                else:
                    latest_year = year
                    try:
                        rates.append\
                    (int_ex/\
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
        tx_prv = self.get_IS_item("Provision for Income Tax")
        ebt = self.get_IS_item("Pretax Income")
        for year in range(6,11):
            if(tx_prv[year]<0):
                yearly.append((-tx_prv[year]) /ebt[year])
            else:
                yearly.append((tx_prv[year])/ebt[year])
        print("Tax Rate: {}".format(np.average(yearly)))
        self.value_drivers_dict["Tax Rate"] = round(np.average(yearly), 4)

    def calculate_ONWC(self):
        yearly = []
        for year in range(6,11):
            currentAssets = self.get_BS_item("Total Current Assets", year)
            cashEquiv = self.get_BS_item("Cash, Cash Equivalents and Short Term Investments", year)
            tcl = self.get_BS_item("Total Current Liabilities", year)
            currentDebt = self.get_BS_item("Current Debt",year)
            rev = self.get_IS_item("Business Revenue", year)
            yearly.append(((currentAssets - cashEquiv) - (tcl-currentDebt))/ rev)
            if(year == 10):
               self.value_drivers_dict["Latest ONWC"] = (currentAssets - cashEquiv) - (tcl-currentDebt)
        print("Operating Net Working Capital to Sales: {}".\
        format(np.average(yearly)))
        self.value_drivers_dict["ONWC to Sales"] = round(np.average(yearly), 4)
        
    def calculate_latest_debt(self):
        try:
            lt = self.BS[self.sheet_index["BS"]["Long Term Debt"]][-1]
            print("Latest Long-term Debt: {}".format(lt))
            self.value_drivers_dict["Latest Long-term debt"] = round(lt, 4)
        except KeyError:
            print("Latest Long term Debt: 0")

    def calculate_shares_outstanding(self):
        '''Market Cap / share price '''
        self.value_drivers_dict["Shares outstanding"] = self.get_IS_item("Diluted Weighted Average Shares Outstanding", -1)
        price = float(self.raw_stock_data["Open"])
        market_cap = self.get_IS_item("Diluted Weighted Average Shares Outstanding", -1) *price

        print("Shares Outstanding : {}".\
        format(self.value_drivers_dict["Shares outstanding"]))

    def calculate_cost_of_equity(self):
        '''CAPM baby.Cost of Equity=risk_free+(Beta*(Market_returns-risk_free))
        '''
        try:
            b = self.raw_stock_data["Beta"] != "N/A"
        except:
            b = "N/A"
        if(b != "N/A"):
            self.value_drivers_dict["Beta"] = float(self.raw_stock_data["Beta"])
            self.value_drivers_dict["Cost of equity"] = \
            self.value_drivers_dict["Risk free rate"] + \
            (self.value_drivers_dict["Beta"] * \
            (.12 - self.value_drivers_dict["Risk free rate"]))

            print("Cost of Equity: {}".format\
            (self.value_drivers_dict["Cost of equity"]))
        else:
            print("Do Gordon Growth Formula")


    def calculate_WACC(self):
        '''(Cost of debt *((Debt / (Equity + Debt)) * (1-Tax rate)))
        + (Cost of equity)*((equity)/(equity + Debt))'''
        try:
            current_price = self.raw_stock_data["Open"]
            self.value_drivers_dict["Current Share Price"] = self.raw_stock_data["Open"]
            cost_debt = self.value_drivers_dict["Cost of debt"]
            cost_equity = self.value_drivers_dict["Cost of equity"]
            lt_debt = self.get_BS_item("Long Term Debt", -1)
            tax_rate = self.value_drivers_dict["Tax Rate"]
            #convert shares outstanding to millions to match financial statement
            #data format
            #total_equity =(self.value_drivers_dict["Shares outstanding"]/1000000)*\
            total_equity = (self.value_drivers_dict["Shares outstanding"]) *\
                current_price

            self.value_drivers_dict["WACC"] = \
            (cost_debt *((lt_debt / (total_equity + lt_debt)) \
            * (1-tax_rate))) + (cost_equity)\
            *((total_equity)/(total_equity + lt_debt))
            print("WACC : {}".format(self.value_drivers_dict["WACC"]))
        except:
            #Company price not available
            print("WACC Can't be calculated")
            self.value_drivers_dict["Current Share Price"] = 0
            self.value_drivers_dict["WACC"] = 0

    def calculate_dividend_rate(self):
        dividend_rate = []
        dividend_growth = []
        last_div = 0
        for y,year in enumerate(self.years,6):
            div = self.get_CF_item("Cash Dividends Paid", y, make_positive = True)
            ni = self.get_IS_item("Net Income after Extraordinary Items and Discontinued Operations", y)

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
        self.value_drivers_dict["Dividend Growth Rate"] = \
        np.average(dividend_growth)
        #for year in range(6,10):

    def calculate_average_stock_repurchase(self):
        stock_repurchase = []
        for y,year in enumerate(self.years,1):
            spent = \
            self.get_CF_item("Payments for Common Stock",y,make_positive = True)
            stock_repurchase.append(spent)
        self.value_drivers_dict["Annual Stock Repurchase"] = \
            np.average(stock_repurchase)



    def value_driver(self, value_driver_name):
        try:
            return(self.value_drivers_dict[value_driver_name])
        except KeyError:
            print("Sorry that isn't in the value drivers list")

    def write_to_csv(self):
        with open\
        ("{}/ValueDrivers.csv".format(self.ticker, self.ticker), "w") as f:
            writer = csv.writer(f, lineterminator='\n')

            for driver in self.value_drivers_list:
                writer.writerow([driver, self.value_drivers_dict[driver]])
