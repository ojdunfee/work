import pandas as pd
import numpy as np
import pickle
import os

class Sheet:
    """Class for cash receipt spreadsheet"""

    def __init__(self, cash='sheets/cash.xls', fees='sheets/fees.xlsx'):
        self.cash = cash
        self.fees = fees
        self.branches = self.load_data('/data/workspace_files/databases/branches.pickle')
        self.accts = self.load_data('/data/workspace_files/databases/accounts.pickle')
        self.accounts = self.load_data('/data/workspace_files/databases/closing.pickle')
        self.df = self.clean_data()
        self.date = self.get_posting_date()
        self.filename = self.get_filename()

    def load_data(self, file):
        """Load dictionaries from pickle file."""
        with open(file, 'rb') as f:
            return pickle.load(f)

    def get_posting_date(self):
        """Get the posting date from the spreadsheet."""
        return self.df['PaymentDate'][0]

    def get_filename(self):
        """
        Get the filename for the current posting, make the directory for accounting sheet if
        it doesnt exist already.
        """
        os.makedirs('cash_receipts/', exist_ok=True)
        return 'cash_receipts/{} cash_receipts.xlsx'.format(self.date.replace('/', '_'))

    def clean_data(self):
        """
        Clean the spreadsheet data, correct errors, replace values, generate columns for company branches and state codes,
        convert datetime to string date, drop rows with $0 values.
        """
        df = pd.read_excel(self.cash, converters={'AcctCode': str, 'Invoice Line Total': float})
        df.State.replace(np.nan, 'NJ', inplace=True)
        df.AcctCode.replace('40003', '40000', inplace=True)
        df.OrderCategory.replace(25, 8, inplace=True)
        df.dropna(subset=['Invoice Line Total'], inplace=True)
        
        df.loc[df[(df['AcctCode'].isin(['40000']) & (~df.OrderCategory.isin([1,4])))].index.tolist(), 'OrderCategory'] = 1
        df.loc[df[(df['AcctCode'].isin(['40002']) & (~df.OrderCategory.isin([2,5])))].index.tolist(), 'OrderCategory'] = 2
        df.loc[df[(df.AcctCode == '40000') & (df.OrderCategory.isin([2,5]))].index.tolist(), 'AcctCode'] = '40002'

        df['File Number'].replace('(CS\d{3})', 'ST-01', regex=True, inplace=True)
        df['File Number'].replace('(ID-R)', 'ID-01', regex=True, inplace=True)
        df['File Number'].replace('(SL-R)', 'SL-01', regex=True, inplace=True)
        df['File Number'].replace('(ID-R)', 'ID-01', regex=True, inplace=True)

        df['File'] = df['File Number'].str[-5:]

        df['Branch'] = df['File'].str[-5:].map(lambda x: self.branches[x] if x in self.branches.keys() else '000')
        df['state_code'] = df['File'].str[-2:]
        df['dept'] = ['02' if x.startswith('6') else '00' for x in df.AcctCode]
        df['PaymentDate'] = pd.to_datetime(df['PaymentDate']).dt.strftime('%m/%d/%Y')
        df = df[df['Invoice Line Total'] != 0]
        return df

class Cash(Sheet):

    def __init__(self, escrow, df):
        super().__init__()
        self.sheet = self.accts[escrow]['sheet']
        self.report = self.get_entries(escrow, df)

    def get_entries(self, escrow, df):
        """
        Generates the general ledger grouping by company -> branch -> state
        Accounts for shortage postings keeping them separate from grouping tagging the file number and agent name
        for further review from Senior Accountants.
        """
        shorts = ['66300', '66302']
        cash = df[~df.AcctCode.isin(shorts)]
        shortages = df[df.AcctCode.isin(shorts)]
        cash = cash.groupby(['TitleCoNum','Branch','state_code','AcctCode']).agg({'Invoice Line Total':'sum',
                                                                                      'PaymentDate': 'first',
                                                                                      'dept': 'first',
                                                                                      'File Number':'first',
                                                                                      'CloseAgent':'first',
                                                                                      'File': 'first'}).reset_index()
        cash = pd.concat([cash, shortages], ignore_index=True)

        cash['Type'] = ['G/L Account' for _ in range(len(cash))]
        cash['Account Desr'] = ['{} {}'.format(cash['File Number'][i], cash['CloseAgent'][i]) if cash.AcctCode[i] in shorts else np.nan for i in range(len(cash))]
        cash['Description Reference'] = ['{} RQ DEP'.format(cash['PaymentDate'][i]) for i in range(len(cash))]
        cash['Debits'] = [abs(round(x, 2)) if x < 0 else np.nan for x in cash['Invoice Line Total']]
        cash['Credits'] = [round(x, 2) if x >= 0 else np.nan for x in cash['Invoice Line Total']]
        cash = cash[['PaymentDate', 'Type', 'AcctCode', 'state_code', 'Branch', 'dept', 'Account Desr', 'Description Reference','Debits', 'Credits']]
        cash.rename(columns={'PaymentDate': 'Date', 'AcctCode': 'Account', 'state_code': 'St',
                             'dept': 'Dept'}, inplace=True)
        cash.sort_values(by=['Branch', 'Account'], inplace=True)
            
        totals = pd.DataFrame({
            'Date': [self.date],
            'Type': ['Bank Account'],
            'Account': [self.accts[escrow]['bank']],
            'St': ['00'],
            'Branch': ['000'],
            'Dept': ['00'],
            'Account Desr': [np.nan],
            'Description Reference': ['{} RQ DEP'.format(self.date)],
            'Debits': [round(df['Invoice Line Total'].sum(), 2)],
            'Credits': [np.nan],
            })
        return pd.concat([cash, totals], ignore_index=True)

class Counts(Sheet):
    """Class for the second half of double sided entries."""

    def __init__(self, escrow, df):
        super().__init__()
        self.frame = df
        self.df = self.group_data(df)
        self.debits = list(self.debits())
        self.report = self.report()
        

    def group_data(self, df):
        """
        groups the dataframe instiniated at Counts to accumulate revenue type totals and counts
        """
        return df.groupby(['TitleCoNum','state_code','OrderCategory']).agg({'Invoice Line Total':'sum',
                                                                            'File':'first', 'Branch':'first',
                                                                            'dept':'first'}).reset_index()

    def debits(self):
        """
        Generates one column of all totals including revenue dollar amounts and sale counts
        """
        totals = self.df['Invoice Line Total'].tolist()
        counts = []
        for (tco, st, oc), df in self.frame.groupby(['TitleCoNum','state_code','OrderCategory']):
            if oc in [1,4]:
                counts.append(len(set(df[df.AcctCode == '40000']['File Number'])))
            elif oc in [2,5]:
                counts.append(len(set(df[df.AcctCode == '40002']['File Number'])))
            else:
                counts.append(len(set(df['File Number'])))
        for i, cell in enumerate(totals):
            yield round(cell, 2)
            yield counts[i]
        yield np.nan

    def get_accounts(self):
        """
        Yields the proper account for posting based on the OrderCategory column
        """
        for i in range(len(self.df)):
            yield self.accounts[self.df['OrderCategory'][i]]['revenue']
            yield self.accounts[self.df['OrderCategory'][i]]['count']
        yield '99998'
    
    def states(self):
        """
        Yields state codes for posting based on the codes generated from Sheet class
        """
        for i in range(len(self.df)):
            yield self.df['state_code'][i]
            yield self.df['state_code'][i]
        yield '00'

    def co_branches(self):
        """
        Yields company branch twice one for each revenue and closing accounts for posting
        """
        for i in range(len(self.df)):
            yield self.df['Branch'][i]
            yield self.df['Branch'][i]
        yield '000'

    def report(self):
        """
        DataFrame of accounting journal entries, drops blank account values subbed with ????
        """
        df = pd.DataFrame({
            'Date': [self.date] * len(self.debits),
            'Type': ['G/L Account'] * len(self.debits),
            'Account': list(self.get_accounts()),
            'St': list(self.states()),
            'Branch': list(self.co_branches()),
            'Dept': ['00'] * len(self.debits),
            'Account Desr': [np.nan] * len(self.debits),
            'Description Reference': ['{} RQ DEP'.format(self.date)] * len(self.debits),
            'Debits': self.debits,
            'Credits': [np.nan] * len(self.debits),
        })
        df = df[~df.Account.str.startswith('?')]
        return df


def check_sheet(cash, fees='sheets/fee_master.xlsx', master=True):
    """
    Scans the worksheet and returns errors if any 
    """
    errors, revisions = [],[]
    fees = pd.read_excel(fees, usecols='A:C')
    fees.columns = ['Date', 'Amount', 'File']

    fee_files = set(fees['File'].tolist())

    second_invoices = cash[(cash.SortField == 2) & (cash.AcctCode.isin(['40000','40002']))]
    for x in second_invoices['File Number'].unique():
        revisions.append('{} is a second invoice with a premium, review'.format(x))

    if cash.OrderCategory.isna().sum() > 0:
        no_order_category = cash[cash.OrderCategory.isna()]['File Number'].unique()
        for x in no_order_category:
            errors.append('{} has no order category'.format(x))

    if master:
        for file in fee_files:
            posted = cash[cash['File Number'] == file]['Invoice Line Total'].sum()
            transferred = fees[fees['File'] == file]['Amount'].sum()
            if not round(posted, 2) == round(transferred, 2):
                errors.append(f'{file} is off {round(transferred - posted, 2)}')
    return errors, revisions

def create_spreadsheet(filename, arr):
    """
    Takes in the filename used to save the excel workbook and the array of data used to populate the sheets
    creates the final workbook generating a new sheet for each company to post
    """
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        print('Creating', filename.split('/')[-1] + '...')
        for i in range(len(arr)):
            print('Adding sheet', arr[i][1] + '...')
            arr[i][0].to_excel(writer, sheet_name=arr[i][1], index=False)
            workbook = writer.book
            worksheet = writer.sheets[arr[i][1]]
            num_format = workbook.add_format({'num_format': '##0.00'})
            worksheet.set_column(0, 2, 12)
            worksheet.set_column(3, 3, 3)
            worksheet.set_column(4, 5, 7)
            worksheet.set_column(6, 6, 15)
            worksheet.set_column(7, 7, 21)
            worksheet.set_column('I:J', 9, cell_format=num_format)
            worksheet.write_formula('J{}'.format(arr[i][-1] + 1), arr[i][-2])
        print('Worksheet finished')

def report_data(escrow, frame):
    """
    Input: spreadsheet data, and escrow account number.
    Output: tuple of values including dataframe, sheetname, excel formula, and end of sheet index.
    """
    cash = Cash(escrow, frame)
    count = Counts(escrow, frame)
    report = pd.concat([cash.report, count.report], ignore_index=True)
    s = report[report.Type == 'Bank Account'].index.values[0] + 3
    e = report[report.Account == '99998'].index.values[0] + 1
    f = '=SUM(I{}:I{})'.format(s,e)
    return report, cash.sheet, f, e

def fix_accounts(escrow, df):
    """Fix Account numbers for specific companies before creating sheet."""
    if escrow == 146:
        df.loc[df[(df.AcctCode == '96021') & (df.EscrowBank == 146)].index.tolist(), 'AcctCode'] = '96024'
    elif escrow == 219:
        df.loc[df[(df.AcctCode == '96023') & (df.EscrowBank == 219)].index.tolist(), 'AcctCode'] = '96020'
        df.loc[df[(df.AcctCode == '43502') & (df.EscrowBank == 219)].index.tolist(), 'AcctCode'] = '43501'
    return df

def main():
    s = Sheet()
    errors, revisions = check_sheet(s.df, master=False)
    if errors:
        for error in errors:
            print(error)
        return None
    for revision in revisions:
        print(revision)
    print()
    arr = []
    for escrow, frame in s.df.groupby('EscrowBank'):
        frame = fix_accounts(escrow, frame)
        data = report_data(escrow, frame)
        arr.append(data)
    
    arr = sorted(arr, key=lambda x: x[1])

    create_spreadsheet(s.filename, arr)

if __name__ == '__main__':
    main()
