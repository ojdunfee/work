import pandas as pd
import numpy as np
from string import punctuation
import pickle
import calendar
import os

class TD:
    """Class representing a bank statement from TD Bank."""
    def __init__(self, path):
        self.path = path
        self._accts = self.load_data('/data/workspace_files/databases/accts.pickle')
        self._cos = self.load_data('/data/workspace_files/databases/cos.pickle')
        self._employee = self.load_data('/data/workspace_files/databases/employee.pickle')
        self._keywords = self.load_data('/data/workspace_files/databases/keywords.pickle')
        self.df = self.load_df()
        self.date = self.get_posting_date()
        self.report = pd.DataFrame({
                                    'Type': list(self.types()),
                                    'No': list(self.accounts()),
                                    'State': list(self.states()),
                                    'Branch Code': list(self.branches()),
                                    'Dept Code': list(self.depts()),
                                    'Description/Comment': list(self.descriptions()),
                                    'Quantity': list(self.quantity()),
                                    'Direct Unit Cost': list(self.costs()),
                                    'IC Partner Ref Type': list(self.ref_type()),
                                    'IC Partner Code': list(self.ic_partner_codes()),
                                    'IC Partner Reference': list(self.ic_partner_refs()),
        })

    def load_data(self, filename):
        """Load dictionaries from pickle files."""
        with open(filename, 'rb') as f:
            return pickle.load(f)

    def load_df(self):
        """Load the bank statement, drop auto payment charges made to account."""
        df = pd.read_excel(self.path, converters={'MCC/SIC Code':str, 'Originating Account Number':str})
        df = df[df['Merchant Name'] != 'AUTO PAYMENT DEDUCTION']
        return df
    
    def accounts(self):
        """Generate Accounting codes based off of the MCC/SIC Code provided by TD."""
        for i, cell in enumerate(self.df['MCC/SIC Code']):
            if cell in self._accts.keys(): 
                yield self._accts[cell]
            else: 
                yield np.nan

    def types(self):
        """Yield G/L Account values for each entry in the dataframe."""
        for _ in range(len(self.df)): 
            yield 'G/L Account'

    def states(self):
        """Yield state codes based off last 4 digits of account number and branch that account is linked to."""
        for i, cell in enumerate(self.df['Originating Account Number']):
            k = str(cell)[-4:]
            if k in self._employee.keys(): 
                yield self._employee[k]['state']
            else: 
                yield np.nan

    def ref_type(self):
        """Yield G/L Account values for each entry in the dataframe."""
        for i, cell in enumerate(self.df['Account Number']):
            yield 'G/L Account'

    def branches(self):
        """Yield company branch code based on the last 4 digits of the card number."""
        for i, cell in enumerate(self.df['Originating Account Number']):
            k = str(cell)[-4:]
            if k in self._employee.keys(): 
                yield self._employee[k]['branch']
            else: 
                yield np.nan

    def descriptions(self):
        """Get the descriptions from the bank statement, combine with initials of the purchasing agent."""
        L = [cell for i, cell in enumerate(self.df['Merchant Name'])]
        for i, cell in enumerate(self.df['Originating Account Name']):
            if str(cell) == 'nan': 
                continue
            else:
                if len(cell.split()) == 2 and cell != 'COMMERCIAL DEPARTMENT':
                    name = cell.split()
                    initials = name[0][0] + name[1][0]
                    L[i] = '{}-{}'.format(initials, L[i])
                else: 
                    continue
        for descr in L: 
            yield descr

    def quantity(self):
        """Yield an int 1 value for each row in the bank statement."""
        for _ in range(len(self.df)): 
            yield 1

    def costs(self):
        """List amounts provided by the TD Bank Statement."""
        for i, cell in enumerate(self.df['Original Amount']): 
            yield round(cell, 2)

    def ic_partner_codes(self):
        """Yield intercompany codes for purchases made from a partnered branch."""
        for i, cell in enumerate(self.df['Originating Account Number']):
            k = str(cell)[-4:]
            if k in self._employee.keys():
                yield self._employee[k]['ic_code']
            else:
                yield np.nan

    def ic_partner_refs(self):
        """Yield blank values for the length of the spreadsheet."""
        for _ in range(len(self.df)):
            yield np.nan

    def depts(self):
        """Yield proper department code determined by the last 4 digits of the card number and corresponding agent."""
        # create conversion table based on accounting code / employee
        for i, cell in enumerate(self.df['Originating Account Number']): 
            k = str(cell)[-4:]
            if k in self._employee.keys():
                yield self._employee[k]['dept']
            else:
                yield np.nan

    def get_posting_date(self):
        """Get the posting date, the last day of the month the bank statement is accounting for."""
        date = self.df.iloc[0, 0]
        return '{}/{}/{}'.format(date.month, calendar.monthrange(date.year, date.month)[1], date.year)

    def fix_sheet(self):
        """
        Once Account values are generated based off of MCC/SIC Code, scan through the list of company keywords
        replacing values that vary from the translation in previous card statements. Also corrects state and department
        codes for specific accounting code values to fix errors.
        """
        for i, cell in enumerate(self.report['Description/Comment']):
            for char in punctuation:
                if char in cell:
                    cell = cell.replace(char, ' ')
            for word in cell.split():
                for k, v in self._keywords.items():
                    if word in v:
                        self.report.loc[i, 'No'] = k
        self.report.loc[self.report[self.report['Direct Unit Cost'] < 0].index.tolist(), 'No'] = '19999'
        self.report.loc[self.report[self.report['No'] == '63000'].index.tolist(), 'Dept Code'] = '00'
        self.report.loc[self.report[self.report['No'] == '63003'].index.tolist(), 'Dept Code'] = '00'
        self.report.loc[self.report[self.report['Description/Comment'] == 'STANDARD VCF 4.4 100'].index.tolist(), 'No'] = '63004'
        self.report.loc[self.report[self.report['Description/Comment'] == 'STANDARD VCF 4.4 100'].index.tolist(), 'State'] = '00'
        self.report.loc[self.report[self.report['Description/Comment'] == 'STANDARD VCF 4.4 100'].index.tolist(), 'Branch Code'] = '000'
        self.report.loc[self.report[self.report['Description/Comment'] == 'STANDARD VCF 4.4 100'].index.tolist(), 'Dept Code'] = '00'


def generate_td(dir='td_statements'):
    """
    Load each bank statement from the directory, gather the proper sheetname from the statement title,
    run through the TD class and create a new sheet in the workbook for the expenses posting.
    """
    frames = []
    filename = ''
    for file in os.listdir(dir):
        sheet = file.split('TD CARD')[0].strip().split()[0]
        if not dir.endswith('/'): 
            dir += '/'
        td = TD(dir + file)
        td.fix_sheet()
        filename = 'journals/' + TD(dir + file).date.replace('/','_') + ' TD_Statements.xlsx'
        frames.append((td.report, sheet))

    frames = sorted(frames, key=lambda x: x[1])
    
    with pd.ExcelWriter(filename) as writer:
        for i in range(len(frames)): 
            frames[i][0].to_excel(writer, sheet_name=frames[i][1], index=False)

if __name__ == '__main__':
    generate_td()
