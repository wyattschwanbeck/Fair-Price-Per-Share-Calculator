import csv, os
#from openpyxl import Workbook
from xlsxwriter.workbook import Workbook

import pandas as pd

class FPPS_Excel_Injector(object):
    def __init__(self, 
                 calculated_value_drivers, 
                 excel_file, 
                 statement_loc, 
                 years_to_project):

        self.VD = calculated_value_drivers
        self.workbook = Workbook(excel_file)
        self.statements = \
        {"incomeStatement": dict(),\
        "balanceSheet" : dict(), \
        "cashFlow" : dict(), \
        "Projection" : dict(),
        "ValueDrivers" : dict()}

        self.statement_loc = statement_loc

        self.years = \
        ["Projections derived from data downloaded from Morningstar.com"] 
        self.years_to_project = years_to_project
        #Calculated but stored for addition into function frame for later        
        self.cap_ex = []
        
        self.function_frame = []
        self.incomeStatement = pd.read_csv(statement_loc + "incomeStatement.csv")
        
    
    def load_financial_data_to_sheets(self):
        statement_loc = self.statement_loc
        files = ["incomeStatement", "cashFlow", "balanceSheet", "ValueDrivers"]
        for csvfile in files:
            print("Working on importing data to excel for " + csvfile)
            excel_sheet_name = csvfile
            sheet_name = csvfile             
            current_statement = self.get_statement_key(sheet_name)
            #worksheet with csv file name                
            worksheet = self.workbook.add_worksheet(sheet_name)
            #using openpyxl now
            #worksheet = self.workbook.create_sheet(sheet_name)          
            with open(statement_loc +csvfile + ".csv", 'r') as f:
                reader = csv.reader(f)
                row_count = 0
                for r, row in enumerate(reader):
                    if(len(row) == 0):
                        continue
                    column_count = 0
                    row_count += 1                
                    for c, col in enumerate(row):
                        column_count += 1
                        if(c==0):
                            key_name = col
                        self.load_statement_meta_data(current_statement,
                                                 excel_sheet_name,
                                                 col,
                                                 key_name,       
                                                 column_count,
                                                 row_count,
                                                 c)                            
                        #check for negative and decimal numbers
                        if("." in col or "-"):
                            try:
                                 #write the csv file content into it
                                worksheet.write(r, c, float(col))
                            except:
                                worksheet.write(r, c, str(col))
                        elif(col.isdigit()):
                            worksheet.write(r,c,int(col))
                        else:
                            worksheet.write(r,c,str(col))
                          
            if(current_statement == "ValueDrivers"):
                #self.workbook.add_vba_project("./vbaProject.bin")
                #worksheet.insert_button("C1", {"macro" : "prettify",
                #                               "caption" : "Format Data",
                #                               "width" : 100,
                #                               "height" : 30})
                
                #insert COE formula to value drivers
                coe = self.cost_of_equity_formula()
                coe_addr = self.get_st_item\
                ("ValueDrivers", "Cost of equity", 1).split("!")[-1]                
                worksheet.write_formula(coe_addr,str(coe[1]))                   
                
                #insert WACC formula to value drivers
                wacc = self.wacc_formula()
                wacc_addr = self.get_st_item\
                ("ValueDrivers", "WACC", 1).split("!")[-1]                      
                worksheet.write_formula(wacc_addr,str(wacc[1]))
                
                projectionRow = r+3
                self.insert_projection_formulas(projectionRow)
                
                for r_2,row in enumerate(self.function_frame,projectionRow):
                    if(row is None):
                        continue
                        
                    for c, col in enumerate(row,1):
                        if(r_2 == projectionRow):
                            worksheet.write(r_2,c,str(col))
                        elif(c == 1):
                            worksheet.write(r_2,c,str(col))
                        else:
                            worksheet.write_formula(r_2, c, str(col))
                    
    def load_statement_meta_data(self,
                            current_statement, excel_sheet_name,item,item_key,
                            col_count,row_count,c):
            if(c == 0):
                self.statements[current_statement].update\
                ({item:[excel_sheet_name+"!A"+str(row_count)]})
            else:
                self.statements[current_statement][item_key].append(\
                excel_sheet_name+"!"\
                +self.convert_num_to_chars(col_count)+str(row_count))
    
    def get_statement_key(self, sheet_name):
        for keys in self.statements.keys():          
            if(keys in sheet_name):          
                return keys
        return "NA"
        
    def get_st_item(self, sheet_name, item, index):
        if(index == False):
            try:
                return self.statements[sheet_name][item]
            except:
                print("Could not retrieve list {} from {}".format\
                (item, sheet_name))
                return "0"
        try:
            return self.statements[sheet_name][item][index]
        except KeyError:
            if(item == "Total operating expenses"):
                try:
                    return self.statements\
                    [sheet_name]["Total costs and expenses"][index]
                except:
                    print("Unable to retreive {} from {} at loc {}".format\
                    (item, sheet_name, index))
                    return "0"
            print("Unable to retreive {} from {} at loc {}".\
            format(item, sheet_name, index))
            return "0"


    def convert_num_to_chars(self, integer):   
        '''takes int and returns column name within excel (A1 notation)'''    
        conversion = ""
        chars = ["","A","B","C","D","E","F","G","H","I","J","K","L","M","N",\
        "O","P","Q","R","S","T","U","V","W","X","Y","Z"]
        if(integer > 26):
            conversion += chars[int(integer/26)]
            conversion += chars[integer%26]
        else:
            conversion += chars[integer]
        return conversion
        
    def save_projection_meta_data(self, item, row, col):
        if(col == 0):
                self.statements["Projection"].update\
                ({item:["A"+str(row)]})
        else:
            self.statements["Projection"][item].append(\
            self.convert_num_to_chars(col)+str(row))
     
    
    def insert_projection_formulas(self, row_num):
        
        #Handle TTM (Trailing Twelve Months) occurrance in reporting to project
        years_to_project = self.years_to_project        
        self.statements["Projection"].update\
        ({"Years":["A"+str(row_num)]})
        self.years.append(self.incomeStatement.columns[-1])
        for years in range(1, int(years_to_project)+1):
            self.save_projection_meta_data("Years", row_num, years)
            self.years.append(int(self.VD.latest_year)+years)
            


        self.function_frame = [self.years]
        #1
        self.function_frame.append(self.project_revenue_formula(row_num +1))
        #2        
        self.function_frame.append\
        (self.project_operating_expenses_formula(row_num+2))
        
        #need to calculate cap ex to project depreciation
        self.cap_ex = self.project_cap_ex_formula(row_num+8)
        #3       
        self.function_frame.append(self.project_depreciation_formula(row_num+3))
        #4        

        #5
        self.function_frame.append(self.project_EBIT_formula(row_num+4))
        
        #6
        self.function_frame.append(self.project_tax_formula(row_num+5))
             
        #7
        self.function_frame.append(self.project_NI_formula(row_num+6))
        #add back depreciaiton for FCFs
        self.function_frame.append(self.project_depreciation_formula(row_num+3))        
               
        #8
        self.function_frame.append(self.cap_ex)        
        
        self.function_frame.append(self.project_ONWC_formula(row_num+7))
              
        #9
        self.function_frame.append(self.project_ONWC_change_formula(row_num+8))
        
        #10
        self.function_frame.append(self.project_FCF_formula(row_num+9))
    
        self.function_frame.append(self.project_terminal_formula(row_num+10)) 
                
        #11
        self.function_frame.append(self.project_total_FCFs(row_num+11))
               
        #12
        self.function_frame.append(self.project_PV_FCFs(row_num + 12))
        
        #13
        self.function_frame.append(self.project_enterprise_value(row_num+13))
        
        #14
        self.function_frame.append(self.project_FPPS(row_num+14))
        
        #15
        self.function_frame.append(self.project_dividends(row_num+15))
        
        #16
        self.function_frame.append(self.project_shares_outstanding(row_num+16))
    
    def cost_of_equity_formula(self):
        '''CAPM baby. 
        Cost of Equity = risk_free + (Beta*(Market_returns - risk_free))
        If beta is 0 ->
        Cost of Equity = \
        (Next Year's Annual Dividend / Current Stock Price) + Dividend Growth Rate'''
        
        coe = ["Cost of Equity"]             
        beta = self.get_st_item("ValueDrivers", "Beta", 1)        
        rf = self.get_st_item("ValueDrivers", "Risk free rate", 1)
        dg = self.get_st_item("ValueDrivers", "Dividend Growth Rate", 1)
        market_return = self.get_st_item("ValueDrivers", "Market Return", 1)        
        share_price = self.get_st_item("ValueDrivers","Current Share Price",1)
        
        coe.append("if({}=0,(D43/{})+{} ,{} + ({}*({} - {})))".\
        format(beta, share_price, dg, rf, beta, market_return, rf))
        return coe
        
    def wacc_formula(self):
        '''(cost_debt *((lt_debt / (total_equity + lt_debt)) * (1-tax_rate))) +
        (cost_equity)*((total_equity)/(total_equity + lt_debt))'''
        wacc = ["WACC"]
        st = self.statements
        cost_debt = st["ValueDrivers"]["Cost of debt"][1]
        try:        
            lt_debt = st["balanceSheet"]["Long Term Debt"][-1]
        except KeyError:
            lt_debt = "0"
        #total_equity = "({}/1000000)*{}".format\
        total_equity = "({})*{}".format\
        (st["ValueDrivers"]["Shares outstanding"][1], 
         st["ValueDrivers"]["Current Share Price"][1])
        tax_rate = st["ValueDrivers"]["Tax Rate"][1]
        cost_equity = st["ValueDrivers"]["Cost of equity"][1]
        
        wacc.append(\
        '({} *(({} / ({} + {})) * (1-{}))) + ({})*(({})/({} + {}))'.format\
        (cost_debt, lt_debt, total_equity, lt_debt, tax_rate, cost_equity, \
        total_equity, total_equity, lt_debt))
        
        return wacc
         
    def get_revenue(self, index):
        if "Business Revenue" in self.statements["incomeStatement"].keys():
            return self.get_st_item("incomeStatement", "Business Revenue", index)
        else:
            return self.get_st_item("incomeStatement", "Total Revenue", index)
    def project_revenue_formula(self, row_start):
        projected_revenue = \
        ["Business Revenue", self.get_revenue(-1)]
        value_driver = self.get_st_item("ValueDrivers","Sales Growth",1)
        
        
        for y,year in enumerate(self.years[2:], 2):
            if(y == 2):
                projected_revenue.append\
                ("{} * (1+{})".format(projected_revenue[y-1], (value_driver)))
            else:
                previous_revenue_loc = "{}{}".format\
                (self.convert_num_to_chars(y+1), row_start+1)
                projected_revenue.append("{} * (1+{})".format\
                (previous_revenue_loc, value_driver))
        return projected_revenue
        
    def project_operating_expenses_formula(self, row_start):
        op_ex = self.get_st_item\
        ("incomeStatement","Cost of Revenue",-1)   
        
        total_expenses = "{}".format(op_ex)      
        projected_expenses = ["Operating Expenses", total_expenses]
        value_driver = self.get_st_item\
        ("ValueDrivers","Operating Expenses to Sales",1)
        for y,year in enumerate(self.years[2:], 2):
            current_col = self.convert_num_to_chars(y+2)            
            projected_expenses.append\
            ("{} * ({})".format(current_col + str(row_start), (value_driver)))
        
        return projected_expenses
        
    def project_cap_ex_formula(self, row_start):
        projected_cap_ex = \
        ["Capital Expenditure", \
        "{}+{}".format(self.get_st_item("cashFlow","Purchase/Sale and Disposal of Property, Plant and Equipment, Net",9), self.get_st_item("cashFlow","Purchase/Sale of Business, Net",9))]
        
        value_driver = self.get_st_item\
        ("ValueDrivers","Capital Expenditure to Sales",1)

        for y,year in enumerate(self.years[2:], 2):
            current_col = self.convert_num_to_chars(y+2)            
            projected_cap_ex.append\
            (("{} * ({})".format(current_col+str(row_start-6),(value_driver))))
        
        return projected_cap_ex
    
    def project_depreciation_formula(self, row_start):
        projected_dep = \
        ["Depreciation", \
        "{}".format\
        (self.get_st_item("cashFlow","Depreciation, Amortization and Depletion, Non-Cash Adjustment",-1))]
        
        value_driver = self.statements["ValueDrivers"]["Depreciation"][1]
        for y,year in enumerate(self.years[2:], 2):
            left_col = self.convert_num_to_chars(y+1)       
            projected_dep.append\
            (("{} + ({} * {})".format\
            (left_col+str(row_start+1),\
            (value_driver),\
            (left_col)+str(row_start+6))))
        
        return projected_dep
        
    def project_EBIT_formula(self, row_start):
        projected_EBIT = \
        ["EBIT"]

        for y,year in enumerate(self.years[1:], 2):
            current_col = self.convert_num_to_chars(y+1)       
            projected_EBIT.append\
            (("{} + {} - {}".format(current_col + str(row_start-2), \
            current_col + str(row_start-1), (current_col)+str(row_start) )))
        
        return projected_EBIT
        
    def project_tax_formula(self, row_start):
        project_taxes = ["Taxes"]
        vd = self.statements["ValueDrivers"]["Tax Rate"][1]
        
        for y,year in enumerate(self.years[1:], 2):
            current_col = self.convert_num_to_chars(y+1)
            project_taxes.append\
            ("If({}>0,{} * {}, 0)".format\
            (current_col + str(row_start), vd, current_col + str(row_start)))
            
        return project_taxes
        
    def project_NI_formula(self, row_start):
        project_NI = ["Net Operating Income"]
        
        for y,year in enumerate(self.years[1:], 1):
            current_col = self.convert_num_to_chars(y+2)
            project_NI.append\
            ("{} - {}".format(current_col + str(row_start-1), \
            current_col + str(row_start)))
            
        return project_NI
        
    def project_ONWC_formula(self, row_start):
        project_ONWC = ["ONWC", \
        self.statements["ValueDrivers"]["Latest ONWC"][1]]

        vd = self.statements["ValueDrivers"]["ONWC to Sales"][1]
        
        for y,year in enumerate(self.years[2:], 2):
            current_col = self.convert_num_to_chars(y+2)
            project_ONWC.append\
            ("{} * {}".format(current_col + str(row_start-5), vd))
            
        return project_ONWC
        
    def project_ONWC_change_formula(self, row_start):
        project_ONWC_change = ["OWNC Change", 0]
        for y,year in enumerate(self.years[2:], 2):
            last_col = self.convert_num_to_chars(y+1)
            current_col = self.convert_num_to_chars(y+2)            
            project_ONWC_change.append\
            ("{} - {}".format(current_col + str(row_start+2),\
            last_col + str(row_start+2)))
            
        return project_ONWC_change
        
    def project_FCF_formula(self, row_start):
        project_FCF = ["Free Cash Flow", 0]
        share_buyback = \
        self.get_st_item("ValueDrivers","Annual Stock Repurchase",1)        
        
        for y,year in enumerate(self.years[2:],2):
            current_col = self.convert_num_to_chars(y+2)
            project_FCF.append\
            ("{} + {} - {} - {} - {}".format(current_col + str(row_start-2), \
            current_col + str(row_start-1), current_col + str(row_start),\
            current_col + str(row_start+2),share_buyback))
            
        return project_FCF
        
    def project_terminal_formula(self, row_start):
        project_term = ["Terminal Value"]
        ltg = self.statements["ValueDrivers"]["Long Term Growth"][1]
        wacc = self.statements["ValueDrivers"]["WACC"][1]
        for y,year in enumerate(self.years[1:],1):
            current_col = self.convert_num_to_chars(y+2)            
            if(y < len(self.years)-1):
                project_term.append(0)
            else:
                project_term.append("={}*(1+{})/({}-{})".format\
                (current_col+str(row_start+2), ltg, wacc, ltg ))
                
        return project_term
    
    def project_total_FCFs(self, row_start):
        project_FCF = ["Total Free Cash Flows", 0]
        
        for y,year in enumerate(self.years[2:],2):
            current_col = self.convert_num_to_chars(y+2)
            project_FCF.append\
            ("{} + {}".format(current_col + str(row_start+1),\
            current_col + str(row_start+2)))
            
        return project_FCF
        
    def project_PV_FCFs(self, row_start):
        project_PV_FCF = ["Present Value Cash Flows"]
        rf = self.get_st_item("ValueDrivers","Risk free rate",1)
        
        for y,year in enumerate(self.years[1:],1):
            end_col = self.convert_num_to_chars(len(self.years[1:])+2)   
            right_col = self.convert_num_to_chars(y+3)
            
            project_PV_FCF.append(\
            "NPV({},{}:{})".format\
            (rf, right_col+str(row_start+2), end_col + str(row_start+2)))
            
        return project_PV_FCF
    
    
    def project_enterprise_value(self, row_start):
        project_ev = ["Enterprise Value"]
        
        lt_debt = self.get_st_item("balanceSheet","Long Term Debt and Capital Lease Obligation",9)
        st_inv = self.get_st_item\
        ("balanceSheet","Current Debt and Capital Lease Obligation",index=False)
        
        if st_inv=="0":
            avg_st_inv = "0"
        else:
            avg_st_inv = "AVERAGE({}:{})".format(st_inv[1], st_inv[-1])  

        cash = self.get_st_item\
        ("balanceSheet","Cash and Cash Equivalents",9)
        
        for y,year in enumerate(self.years[1:],1): 
            current_col = self.convert_num_to_chars(y+2)
            
            project_ev.append(\
            "{} - {} + {} + {} ".format\
            (current_col + str(row_start+2), lt_debt,  avg_st_inv, cash))
            
        return project_ev
     
    def project_FPPS(self, row_start):
        FPPS = ["Fair Price Per Share"]
        #"({}/1000000)".format\        
        outstanding_shares = \
            "({})".format\
            (self.get_st_item("ValueDrivers","Shares outstanding",1))
        
        for y,year in enumerate(self.years[1:],1): 
            current_col = self.convert_num_to_chars(y+2)
            
            FPPS.append(\
            "{} / {} ".format\
            (current_col + str(row_start+2), current_col + str(row_start +5)))
            
        return FPPS
        
    def project_dividends(self, row_start):
        #"({}/1000000)"          
        outstanding_shares = \
            "({})".format\
            (self.get_st_item("ValueDrivers","Shares outstanding",1))
        
        div_rate = self.get_st_item("ValueDrivers","Dividend Rate",1)
        div_growth = self.get_st_item("ValueDrivers","Dividend Growth Rate",1)
        div_paid = self.get_st_item("cashFlow", "Cash Dividends Paid", -1)
        dividends_paid = \
        ["Dividends per share", "-{}/{}".format(div_paid, outstanding_shares)]
        
        for y,year in enumerate(self.years[2:], 2):
            latest_div_rate = "{}*(1+({}*({}-1)))".format(div_rate,div_growth,y)
            current_col = self.convert_num_to_chars(y+2)
            NI_loc = current_col + str(row_start -8)
            
            dividends_paid.append("({}*({}))/{}".format\
            (NI_loc, latest_div_rate, outstanding_shares))
            
        return dividends_paid
        
    def project_shares_outstanding(self, row_start):
        vd = self.get_st_item("ValueDrivers","Shares outstanding",1)       
        sp = self.get_st_item("ValueDrivers", "Current Share Price", 1)
        bb = self.get_st_item("ValueDrivers", "Annual Stock Repurchase", 1)
        #"({}/1000000)".format\         
        outstanding_shares = \
            "({})".format\
            (vd)        
        SO = ["Shares Outstanding", outstanding_shares]
        
        
        for y,year in enumerate(self.years[2:], 2):
            left_col = self.convert_num_to_chars(y+1)                      
            so = "{}-({}/{})".format(left_col + str(row_start+3),bb,sp)
            SO.append(so)
            
        return SO
            
        
        
        
        
