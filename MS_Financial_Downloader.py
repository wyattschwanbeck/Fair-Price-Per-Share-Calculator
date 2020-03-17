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
from bs4 import BeautifulSoup

class FinancialsDownloader(object):
    u"""Downloads financials from http://financials.morningstar.com/
    """

    def __init__(self, table_prefix = u'morningstar_'):
        u"""Constructs the FinancialsDownloader instance.
        :param table_prefix: Prefix of the MySQL tables.
        """
        self._table_prefix = table_prefix

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
        result = {}

        ##########################
        # Error Handling
        ##########################

        # Empty String
        if len(ticker) == 0:
            raise ValueError("You did not enter a ticker symbol.  Please"
                             " try again.")

        for report_type, table_name in [
                (u'is', u'income_statement'),
                (u'bs', u'balance_sheet'),
                (u'cf', u'cash_flow')]:
            frame = self._download(ticker, report_type)
            result[table_name] = frame
            if conn:
                self._upload_frame(
                    frame, ticker, self._table_prefix + table_name, conn)
        if conn:
            self._upload_unit(ticker, self._table_prefix + u'unit', conn)
        result[u'period_range'] = self._period_range
        result[u'fiscal_year_end'] = self._fiscal_year_end
        result[u'currency'] = self._currency
        return result

    def _download(self, ticker, report_type):
        #&rounding=3
        u"""Downloads and returns a pandas.DataFrame corresponding to the
        given Morningstar ticker and the given type of the report.
        :param ticker: Morningstar ticker.
        :param report_type: Type of the report ('is', 'bs', 'cf').
        :return  pandas.DataFrame corresponding to the given Morningstar ticker
        and the given type of the report.&rounding=3
        """
        url = (r'http://financials.morningstar.com/ajax/' +
               r'ReportProcess4HtmlAjax.html?&t=' + ticker +
               r'&region=usa&culture=en-US&cur=USD' +
               r'&reportType=' + report_type + r'&period=12' +
               r'&dataType=A&order=asc&columnYear=5&platform=sal&view=raw')
        with urllib.request.urlopen(url) as response:
            json_text = response.read().decode(u'utf-8')

            ##############################
            # Error Handling
            ##############################

            # Wrong ticker
            if len(json_text)==0:
                raise ValueError("MorningStar cannot find the ticker symbol "
                                 "you entered or it is INVALID. Please try "
                                 "again.")

            json_data = json.loads(json_text)
            result_soup = BeautifulSoup(json_data[u'result'],u'html.parser')
            return self._parse(result_soup)

    def _parse(self, soup):
        u"""Extracts and returns a pandas.DataFrame corresponding to the
        given parsed HTML response from financials.morningstar.com.
        :param soup: Parsed HTML response by BeautifulSoup.
        :return pandas.DataFrame corresponding to the given parsed HTML response
        from financials.morningstar.com.
        """
        # Left node contains the labels.
        left = soup.find(u'div', u'left').div
        # Main node contains the (raw) data.
        main = soup.find(u'div', u'main').find(u'div', u'rf_table')
        year = main.find(u'div', {u'id': u'Year'})
        self._year_ids = [node.attrs[u'id'] for node in year]
        period_month = pd.datetime.strptime(year.div.text, u'%Y-%m').month
        self._period_range = pd.period_range(
            year.div.text, periods=len(self._year_ids),
            # freq=pd.datetools.YearEnd(month=period_month))
            freq = pd.tseries.offsets.YearEnd(month=period_month))
        unit = left.find(u'div', {u'id': u'unitsAndFiscalYear'})
        self._fiscal_year_end = int(unit.attrs[u'fyenumber'])
        self._currency = unit.attrs[u'currency']
        self._data = []
        self._label_index = 0
        self._read_labels(left)
        self._data_index = 0
        self._read_data(main)
        return pd.DataFrame(self._data,
                            #u'parent_index',
                            columns=[ u'title'] + list(
                                self._period_range))

    def _read_labels(self, root_node, parent_label_index = None):
        u"""Recursively reads labels from the parsed HTML response.
        """
        for node in root_node:
            if node.has_attr(u'class') and u'r_content' in node.attrs[u'class']:
                self._read_labels(node, self._label_index - 1)
            if (node.has_attr(u'id') and
                    node.attrs[u'id'].startswith(u'label') and
                    not node.attrs[u'id'].endswith(u'padding') and
                    (not node.has_attr(u'style') or
                        u'display:none' not in node.attrs[u'style'])):
                label_id = node.attrs[u'id'][6:]
                label_title = (node.div.attrs[u'title']
                               if node.div.has_attr(u'title')
                               else node.div.text)
                self._data.append({
                    u'id': label_id,
                    u'index': self._label_index,
                    #u'parent_index': (parent_label_index
                    #                 if parent_label_index is not None
                    #                 else self._label_index),
                    u'title': label_title})
                self._label_index += 1

    def _read_data(self, root_node):
        u"""Recursively reads data from the parsed HTML response.
        """
        for node in root_node:
            if node.has_attr(u'class') and u'r_content' in node.attrs[u'class']:
                self._read_data(node)
            if (node.has_attr(u'id') and
                    node.attrs[u'id'].startswith(u'data') and
                    not node.attrs[u'id'].endswith(u'padding') and
                    (not node.has_attr(u'style') or
                        u'display:none' not in node.attrs[u'style'])):
                data_id = node.attrs[u'id'][5:]
                while (self._data_index < len(self._data) and
                       self._data[self._data_index][u'id'] != data_id):
                    # In some cases we do not have data for all labels.
                    self._data_index += 1
                assert(self._data_index < len(self._data) and
                       self._data[self._data_index][u'id'] == data_id)
                for (i, child) in enumerate(node.children):
                    try:
                        value = float(child.attrs[u'rawvalue'])
                    except ValueError:
                        value = None
                    self._data[self._data_index][
                        self._period_range[i]] = value
                self._data_index += 1
