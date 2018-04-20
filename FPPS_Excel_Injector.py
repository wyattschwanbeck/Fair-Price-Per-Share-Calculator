import csv, os
from xlsxwriter.workbook import Workbook
class FPPS_Excel_Injector(object):
    def __init__(self, calculated_value_drivers, excel_file, statement_loc, years_to_project):
        self.VD = calculated_value_drivers
        self.workbook = Workbook(excel_file)
        self.statements = \
        {"Value Drivers" : dict(), \
        "Income Statement": dict(),\
        "Balance Sheet" : dict(), \
        "Cash Flow" : dict(), \
        "N/A" : dict()}
        self.statement_loc = statement_loc
        self.years = ["Data In Millions"] 
        self.years_to_project = years_to_project
        #Calculated but stored for addition into function frame for later        
        self.cap_ex = []
        
        
        self.function_frame = []
    
    
    def load_financial_data_to_sheets(self):
        statement_loc = self.statement_loc
        for csvfile in os.listdir(statement_loc):
            if(csvfile[-4:] == ".csv"):     
                print("Working on importing data to excel for " + csvfile)
                excel_sheet_name = "'{}'".format(csvfile[0:-4])
                sheet_name = csvfile[0:-4]             
                current_statement = self.get_statement_key(sheet_name)
                worksheet = self.workbook.add_worksheet(sheet_name) #worksheet with csv file name
                with open(statement_loc +csvfile, 'r') as f:
                    reader = csv.reader(f)
                    row_count = 0
                    for r, row in enumerate(reader):
                        if(len(row) == 0):
                            continue
                        column_count = 0
                        row_count += 1                
                        for c, col in enumerate(row):
                            column_count += 1
                            if(c == 0):
                                self.statements[current_statement].update\
                                ({col : [excel_sheet_name + "!A" + str(row_count)]})
                                key_name = col
                            else:
                                self.statements[current_statement][key_name].append(\
                                excel_sheet_name + "!" +self.convert_num_to_chars(column_count) + str(row_count))
                            
                            
                            if("." in col or "-"):
                                try:                
                                    worksheet.write(r, c, float(col)) #write the csv file content into it
                                except:
                                    worksheet.write(r, c, str(col))
                            elif(col.isdigit()):
                                worksheet.write(r,c,int(col))
                            else:
                                worksheet.write(r,c,str(col))
                              
                if(current_statement == "Value Drivers"):
                    #insert COE formula to value drivers
                    coe = self.cost_of_equity_formula()
                    worksheet.write_formula(r-10,1,str(coe[1]))
                    wacc = self.wacc_formula()
                    worksheet.write_formula(r-7,1,str(wacc[1]))
                    self.insert_projection_formulas(r+3)
                    for r_2,row in enumerate(self.function_frame, r+3):
                        if(row is None):
                            continue
                        if(row[0] == "Cost of Equity"):
                            
                            continue
                        for c, col in enumerate(row,1):
                            if(r_2 == r+3):
                                worksheet.write(r_2,c,str(col))
                            elif(c == 1):
                                worksheet.write(r_2,c,str(col))
                            else:
                                worksheet.write_formula(r_2, c, str(col))
                    
    
    def get_statement_key(self, sheet_name):
        for keys in self.statements.keys():          
            if(keys in sheet_name):          
                return keys
        return "N/A"
        
    def get_st_item(self, sheet_name, item, index):
        try:
            return self.statements[sheet_name][item][index]
        except KeyError:
            if(item == "Total operating expenses"):
                try:
                    return self.statements[sheet_name]["Total costs and expenses"][index]
                except:
                    print("Unable to retreive {} from {} at loc {}".format(item, sheet_name, index))
                    return "0"
            print("Unable to retreive {} from {} at loc {}".format(item, sheet_name, index))
            return "0"


    def convert_num_to_chars(self, integer):   
        '''takes int and returns column name within excel (A1 notation)'''    
        conversion = ""
        chars = ["","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P",\
        "Q","R","S","T","U","V","W","X","Y","Z"]
        if(integer > 26):
            conversion += chars[int(integer/26)]
            conversion += chars[integer%26]
        else:
            conversion += chars[integer]
        return conversion
        
      
    def insert_projection_formulas(self, row_num):
        years_to_project = self.years_to_project        
        for years in range(0, int(years_to_project)+1):
            self.years.append(int(self.VD.latest_year)+years)
        

        self.function_frame = [self.years]
        
        #1
        self.function_frame.append(self.project_revenue_formula(row_num +1))
        
        #2        
        self.function_frame.append(self.project_operating_expenses_formula(row_num + 2))
        #need to calculate cap ex to calculate depreciation
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
        
        self.function_frame.append(self.project_ONWC_formula(row_num + 7))
              
        #9
        self.function_frame.append(self.project_ONWC_change_formula(row_num + 8))
        
        #10
        self.function_frame.append(self.project_FCF_formula(row_num+9))
    
        self.function_frame.append(self.project_terminal_formula(row_num+10)) 
                
        #11
        self.function_frame.append(self.project_total_FCFs(row_num +11))
               
        #12
        self.function_frame.append(self.project_PV_FCFs(row_num + 12))
        
        #13
        self.function_frame.append(self.project_enterprise_value(row_num+13))
        
        #14
        self.function_frame.append(self.project_FPPS(row_num+14))
        
        #15
        self.function_frame.append(self.project_dividends(row_num + 15))
        
        #16
        self.function_frame.append(self.project_shares_outstanding(row_num + 16))
    
    def cost_of_equity_formula(self):
        '''CAPM baby. 
        Cost of Equity = risk_free + (Beta*(Market_returns - risk_free))
        If beta is 0 ->
        Cost of Equity = \
        (Next Year's Annual Dividend / Current Stock Price) + Dividend Growth Rate'''
        
        coe = ["Cost of Equity"]             
        beta = self.statements["Value Drivers"]["Beta"][1]
        rf = self.statements["Value Drivers"]["Risk free rate"][1]
        dg = self.statements["Value Drivers"]["Dividend Growth Rate"][1]
        market_return = self.statements["Value Drivers"]["Market Return"][1]
        share_price = self.statements["Value Drivers"]["Current Share Price"][1]
        coe.append("if({}=0,(D43/{})+{} ,{} + ({}*({} - {})))".format(beta, share_price, dg, rf, beta, market_return, rf))
        return coe
        
    def wacc_formula(self):
        '''(cost_debt *((lt_debt / (total_equity + lt_debt)) * (1-tax_rate))) + (cost_equity)*((total_equity)/(total_equity + lt_debt))'''
        wacc = ["WACC"]
        st = self.statements
        cost_debt = st["Value Drivers"]["Cost of debt"][1]
        try:        
            lt_debt = st["Balance Sheet"]["Long-term debt"][-1]
        except KeyError:
            lt_debt = "0"
        #total_equity = "({}/1000000)*{}".format\
        total_equity = "({})*{}".format\
        (st["Value Drivers"]["Shares outstanding"][1], 
         st["Value Drivers"]["Current Share Price"][1])
        tax_rate = st["Value Drivers"]["Tax Rate"][1]
        cost_equity = st["Value Drivers"]["Cost of equity"][1]
        
        wacc.append(\
        '({} *(({} / ({} + {})) * (1-{}))) + ({})*(({})/({} + {}))'.format\
        (cost_debt, lt_debt, total_equity, lt_debt, tax_rate, cost_equity, \
        total_equity, total_equity, lt_debt))
        
        return wacc
         
    
    def project_revenue_formula(self, row_start):
        projected_revenue = ["Revenue", self.statements["Income Statement"]["Revenue"][-1]]
        value_driver = self.statements["Value Drivers"]["Sales Growth"][1]
        for y,year in enumerate(self.years[2:], 2):
            if(y == 2):
                projected_revenue.append("{} * (1+{})".format(projected_revenue[y-1], (value_driver)))
            else:
                previous_revenue_loc = "{}{}".format\
                (self.convert_num_to_chars(y+1), row_start+1)
                projected_revenue.append("{} * (1+{})".format\
                (previous_revenue_loc, value_driver))
        return projected_revenue
        
    def project_operating_expenses_formula(self, row_start):
        op_ex = self.get_st_item("Income Statement", "Total operating expenses", -1)        
        total_cor = self.get_st_item("Income Statement", "Cost of revenue", -1)
        total_expenses = "{} + {}".format(op_ex, total_cor)      
        projected_expenses = ["Operating Expenses", total_expenses]
        value_driver = self.statements["Value Drivers"]["Operating Expenses to Sales"][1]
        for y,year in enumerate(self.years[2:], 2):
            current_col = self.convert_num_to_chars(y+2)            
            projected_expenses.append\
            ("{} * ({})".format(current_col + str(row_start), (value_driver)))
        
        return projected_expenses
        
    def project_cap_ex_formula(self, row_start):
        projected_cap_ex = \
        ["Capital Expenditure", \
        "-{}".format(self.statements["Cash Flow"]["Capital expenditure"][-1])]
        
        value_driver = self.statements["Value Drivers"]["Capital Expenditure to Sales"][1]
        for y,year in enumerate(self.years[2:], 2):
            current_col = self.convert_num_to_chars(y+2)            
            projected_cap_ex.append\
            (("{} * ({})".format(current_col + str(row_start-6), (value_driver))))
        
        return projected_cap_ex
    
    def project_depreciation_formula(self, row_start):
        projected_dep = \
        ["Depreciation", \
        "{}".format(self.statements["Cash Flow"]["Depreciation & amortization"][-1])]
        
        value_driver = self.statements["Value Drivers"]["Depreciation"][1]
        for y,year in enumerate(self.years[2:], 2):
            left_col = self.convert_num_to_chars(y+1)       
            projected_dep.append\
            (("{} + ({} * {})".format(left_col + str(row_start+1), (value_driver), (left_col)+str(row_start +6) )))
        
        return projected_dep
        
    def project_EBIT_formula(self, row_start):
        projected_EBIT = \
        ["EBIT"]

        for y,year in enumerate(self.years[1:], 1):
            current_col = self.convert_num_to_chars(y+2)       
            projected_EBIT.append\
            (("{} - {} - {}".format(current_col + str(row_start-2), \
            current_col + str(row_start-1), (current_col)+str(row_start) )))
        
        return projected_EBIT
        
    def project_tax_formula(self, row_start):
        project_taxes = ["Taxes"]
        vd = self.statements["Value Drivers"]["Tax Rate"][1]
        
        for y,year in enumerate(self.years[1:], 1):
            current_col = self.convert_num_to_chars(y+2)
            project_taxes.append\
            ("If({}>0,{} * {}, 0)".format(current_col + str(row_start), vd, current_col + str(row_start)))
            
        return project_taxes
        
    def project_NI_formula(self, row_start):
        project_NI = ["Net Income"]
        
        for y,year in enumerate(self.years[1:], 1):
            current_col = self.convert_num_to_chars(y+2)
            project_NI.append\
            ("{} - {}".format(current_col + str(row_start-1), \
            current_col + str(row_start)))
            
        return project_NI
        
    def project_ONWC_formula(self, row_start):
        project_ONWC = ["ONWC", \
        self.statements["Value Drivers"]["Latest ONWC"][1]]

        vd = self.statements["Value Drivers"]["ONWC"][1]
        
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
        
        for y,year in enumerate(self.years[2:],2):
            current_col = self.convert_num_to_chars(y+2)
            project_FCF.append\
            ("{} + {} - {} - {}".format(current_col + str(row_start-2), \
            current_col + str(row_start-1), current_col + str(row_start),\
            current_col + str(row_start+2)))
            
        return project_FCF
        
    def project_terminal_formula(self, row_start):
        project_term = ["Terminal Value"]
        ltg = self.statements["Value Drivers"]["Long Term Growth"][1]
        wacc = self.statements["Value Drivers"]["WACC"][1]
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
        rf = self.statements["Value Drivers"]["Risk free rate"][1]
        
        for y,year in enumerate(self.years[1:],1):
            end_col = self.convert_num_to_chars(len(self.years[1:])+2)   
            right_col = self.convert_num_to_chars(y+3)
            
            project_PV_FCF.append(\
            "NPV({},{}:{})".format\
            (rf, right_col+str(row_start+2), end_col + str(row_start+2)))
            
        return project_PV_FCF
    
    
    def project_enterprise_value(self, row_start):
        project_ev = ["Enterprise Value"]
        try:        
            lt_debt = self.statements["Balance Sheet"]["Long-term debt"][-1]
        except KeyError:
            lt_debt = "0"
        try:
            st_inv = self.statements["Balance Sheet"]["Short-term investments"]
            avg_st_inv = "AVERAGE({}:{})".format(st_inv[1], st_inv[-1])        
        except:
            avg_st_inv = 0
        cash = self.statements["Balance Sheet"]["Cash and cash equivalents"][-1]  
        
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
            (self.statements["Value Drivers"]["Shares outstanding"][1])
        
        for y,year in enumerate(self.years[1:],1): 
            current_col = self.convert_num_to_chars(y+2)
            
            FPPS.append(\
            "{} / {} ".format\
            (current_col + str(row_start+1), current_col + str(row_start +5)))
            
        return FPPS
        
    def project_dividends(self, row_start):
        #"({}/1000000)"          
        outstanding_shares = \
            "({})".format\
            (self.statements["Value Drivers"]["Shares outstanding"][1])
        
        div_rate = self.statements["Value Drivers"]["Dividend Rate"][1]
        div_growth = self.get_st_item("Value Drivers","Dividend Growth Rate",1)
        div_paid = self.get_st_item("Cash Flow", "Dividend paid", 1)
        dividends_paid = \
        ["Dividends per share", "-{}/{}".format(div_paid, outstanding_shares)]
        
        for y,year in enumerate(self.years[2:], 2):
            latest_div_rate = "{}*(1+({}*({}-1)))".format(div_rate,div_growth,y)
            current_col = self.convert_num_to_chars(y+2)
            NI_loc = current_col + str(row_start -8)
            
            dividends_paid.append("({}*({}))/{}".format(NI_loc, latest_div_rate, outstanding_shares))
            
        return dividends_paid
        
    def project_shares_outstanding(self, row_start):
        vd = self.statements["Value Drivers"]["Shares outstanding"][1]       
        sp = self.get_st_item("Value Drivers", "Current Share Price", 1)
        bb = self.get_st_item("Value Drivers", "Annual Stock Repurchase", 1)
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
