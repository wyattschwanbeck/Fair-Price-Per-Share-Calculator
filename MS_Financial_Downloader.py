# Copyright (c) 2015 Wyatt Schwanbeck
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
# THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
"""Module for downloading financial data from financials.morningstar.com.
"""

import json
import pandas as pd
import urllib.request
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup
import sqlite3
import os
import csv

class FinancialsDownloader(object):
    u"""Downloads financials from http://financials.morningstar.com/
    """

    def __init__(self, table_prefix = u'morningstar_'):
        u"""Constructs the FinancialsDownloader instance.
        :param table_prefix: Prefix of the MySQL tables.
        """
        self._table_prefix = table_prefix
        self.conn = sqlite3.connect("tickers.db")
        self.c = self.conn.cursor()
    def __del__(self):
        self.conn.close()

    def download(self, ticker, conn = None):
        u"""Downloads and returns a dictionary containing pandas.DataFrames
        representing the financials (i.e. income statement, balance sheet,
        cash flow) for the given Morningstar ticker. If the MySQL connection
        is specified then the downloaded financials are uploaded to the MySQL
        database.
        :param ticker: Morningstar ticker.
        :param conn: MySQL connection.
        :return Dictionary containing pandas.DataFrames representing the
        financials for the given Morningstar ticker.
        """
        reportTypes = ["incomeStatement", "balanceSheet","cashFlow"]
        result = {}
        
        self.c.execute("SELECT tickerSymbol, performanceId, exchangeLabel FROM Tickers where tickerSymbol = ?;", [ticker.upper()])
        result = [row[0:3] for row in self.c.fetchall()]
        print(result)
        performanceId = result[0][1]
        exchange = result[0][2]
        for reportType in reportTypes:
            url = f"https://api-global.morningstar.com/sal-service/v1/stock/newfinancials/{performanceId}/{reportType}/detail?dataType=A&reportType=A&locale=en&languageId=en&locale=en&clientId=MDC&component=sal-equity-financials-details&version=4.30.0"
            req = Request(url)
            req.add_header("Referer",f"https://www.morningstar.com/stocks/{exchange}/{ticker}/financials")
            req.add_header("apikey", 'lstzFDEOhfFNMLikKa0am9mgEKLBl49T')

            content = urlopen(req).read()
            data = json.loads(content.decode('utf-8'))
            newpath = "{}/".format(ticker)     
            if not os.path.exists(newpath):
                os.makedirs(newpath)
                print("Creating folder for {}".format(ticker))
            with open(f"{ticker}/{reportType}.csv", 'w', newline='') as csvfile:
                def recursiveSublevel(subLevels):
                    for subLevel in subLevels:
                        if "datum" in subLevel.keys():
                            if "eps" not in subLevel['label'].lower() and " per " not in subLevel["label"].lower():
                                
                                row_values= [subLevel['label']] + [0 if value =="_PO_" or value is None else str(value * orderOfMagnitude) for value in subLevel["datum"]]
                            else: 
                                row_values = [subLevel['label']] + [0 if value =="_PO_" or value is None else str(value) for value in subLevel["datum"]]
                            writer.writerow(row_values)
                        if "subLevel" in subLevel.keys():
                            recursiveSublevel(subLevel["subLevel"])
                        
                writer = csv.writer(csvfile)
                header_row = [''] + data['columnDefs']
                writer.writerow(header_row)
                adjusted = data['footer']['orderOfMagnitude']
                orderOfMagnitude = 0
                if adjusted == "Billion":
                    orderOfMagnitude = 1000000000
                elif adjusted == "Million":
                    orderOfMagnitude = 1000000
                elif adjusted =="Thousand":
                    orderOfMagnitude = 1000
                for row in data["rows"]:
                    #Parse data via datum key
                    if "datum" in row.keys():
                        

                        if "eps" not in row['label'].lower() and " per " not in row["label"].lower():
                            
                            row_values= [row['label']] + [0 if value =="_PO_" or value is None else str(value * orderOfMagnitude) for value in row["datum"]]
                        else: 
                            row_values = [row['label']] + [0 if value =="_PO_" or value is None else str(value) for value in row["datum"]]
                        writer.writerow(row_values)
                    if "subLevel" in row.keys():
                        recursiveSublevel(row["subLevel"])
                