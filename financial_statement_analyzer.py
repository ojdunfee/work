import pandas as pd
import numpy as np
from dateutil import parser
import pickle

def get_dates(path):
    """
    load financial statement into pandas. grab the start and end dates from the worksheet to filter accounting data.
    """
    flags = ['Start Date:', 'End Date:']
    xl = pd.ExcelFile(path)
    df = xl.parse('Inc Stmt - CMVPM - Detail')
    df = df.iloc[1:-4, 2:]
    start, end = df[df.iloc[:, 1].isin(flags)].values[:, 2]
    return parser.parse(start), parser.parse(end)

def parse_variance(start, statement='data/financial_statement.xlsx', detail='data/detail.xlsx', threshold=500):
    """
    provide the start date, financial statement, and general ledger to scan the financial statement and examine 
    expense differences above a desired threshold. If expenses are above the desired variance/threshold 
    gather general ledger entries sorted by most expensive to least expensive. While the sum of these charges is
    less than the amount of variance print statement notes to copy onto worksheet.
    """
    with open('expense_accounts.pickle', 'rb') as f:
        D = pickle.load(f)
        
    df = pd.read_excel(statement, sheet_name='Inc Stmt - CMVPM - Summary')
    df = df.iloc[2:, 2:]

    arr = []

    for i, cell in enumerate(df.iloc[:, 1]):
        if cell in D.keys():
            if df.iloc[i, 1:].values[-4] >= threshold:
                arr.append((cell, D[cell], round(df.iloc[i, 1:].values[-4], 2)))

    df2 = pd.read_excel(detail, skiprows=list(range(0,10)), usecols='F:Z', header=[1],
                       converters={'Posting Date':pd.to_datetime, 'Acct #':str})

    df2 = df2[df2['Posting Date'] >= start]
    L = []
    for i in range(len(arr)):
        variance = 0
        frame = df2[df2['Acct #'].isin(arr[i][1])].sort_values(by='Debit Amount', ascending=False)
        amts = frame['Debit Amount'].tolist()
        vendors = frame['Vendor Name'].tolist()
        descriptions = frame['G/L Description'].tolist()
        j = 0
        vdict = {arr[i][0]: 'Paid '}
        while variance < arr[i][-1]:
            if not amts[j] == 0:
                if str(vendors[j]) == 'nan':
                    vdict[arr[i][0]] += '${} to {}, '.format(amts[j], descriptions[j].lower().title())
                else:
                    vdict[arr[i][0]] += '${} to {}, '.format(amts[j], vendors[j].lower().title())
            variance += amts[j]
            j += 1

        for k, v in vdict.items():
            s = '''{}) {}\n{}'''.format(str(i+1), arr[i][0], vdict[arr[i][0]].strip(', '))
            yield s

if __name__ == '__main__':
    start, end = get_dates('data/financial_statement.xlsx')
    for x in parse_variance(start=start):
        print(x)
        print()
