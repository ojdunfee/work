import pandas as pd
import numpy as np
import os
import pickle


def load_pickle(file):
    with open(file, 'rb') as f:
        return pickle.load(f)

def get_state(cell):
    if cell.startswith('CS'):
        return '01'
    else:
        if cell.split('-')[-1].isnumeric() and len(cell.split('-')[-1]) == 2:
            return cell.split('-')[-1]
        elif cell.split('-')[-1] == 'R':
            return '01'
        else:
            return cell.split('-')[-1][:-1]

def get_dept(cell):
    if cell.startswith('6'):
        return '02'
    else:
        return '00'

def debits(cell):
    if cell < 0:
        return round(abs(cell), 2)
    else:
        return np.nan

def credits(cell):
    if cell >= 0:
        return round(cell, 2)
    else:
        return np.nan

def report(escrow, df):
    cash = df[~df.AcctCode.isin(['66302','66300'])]
    shorts = df[df.AcctCode.isin(['66302', '66300'])]

    frame = cash.groupby(['TitleCoNum','Branch','St','AcctCode']).agg({
        'Invoice Line Total':'sum',
        'Dept':'first',
        'CloseAgent':'first',
        'Date': 'first'
    }).reset_index()

    df = pd.concat([frame, shorts], ignore_index=True)
    df['Type'] = ['G/L Account'] * len(df)
    df['Account Desr'] = ['{} {}'.format(df['File Number'][i], df['CloseAgent'][i]) if df.AcctCode[i] in ['66300','66302'] else np.nan for i in range(len(df))]
    df['Description Reference'] = ['{} RQ DEP'.format(df['Date'][i]) for i in range(len(df))]
    df['Debits'] = df['Invoice Line Total'].apply(debits)
    df['Credits'] = df['Invoice Line Total'].apply(credits)

    totals = pd.DataFrame({
        'Date': df['Date'][0],
        'Type': ['Bank Account'],
        'AcctCode': [accounts[escrow]['bank']],
        'St': ['00'],
        'Branch': ['000'],
        'Dept': ['00'],
        'Account Desr': [np.nan],
        'Description Reference': ['{} RQ DEP'.format(df['Date'][0])],
        'Debits': [round(df['Invoice Line Total'].sum(), 2)],
        'Credits': [np.nan],
    })

    df = pd.concat([df, totals], ignore_index=True)
    df.rename(columns={'AcctCode':'Account'}, inplace=True)

    return df[['Date','Type','Account','St','Branch','Dept','Account Desr','Description Reference','Debits','Credits']]

def totals(escrow, df):
    debits, credits, states = list(), list(), list()
    accounts, branches = list(), list()
    dates, description_references = list(), list()
    for (tco, date, branch, st, oc), frame in df.groupby(['TitleCoNum', 'Date', 'Branch', 'St','OrderCategory']):
        if oc in [1,4]:
            debits.append(round(frame['Invoice Line Total'].sum(), 2))
            debits.append(len(set(frame[frame.AcctCode == '40000']['File Number'])) - len(set(frame[(frame.SortField == 2) & (frame.AcctCode == '40000')]['File Number'])))
        elif oc in [2,5]:
            debits.append(round(frame['Invoice Line Total'].sum(), 2))
            debits.append(len(set(frame[frame.AcctCode == '40002']['File Number'])) - len(set(frame[(frame.SortField == 2) & (frame.AcctCode == '40002')]['File Number'])))
        else:
            debits.append(round(frame['Invoice Line Total'].sum(), 2))
            debits.append(len(set(frame['File Number'].tolist())))
        accounts.append(closing[oc]['revenue'])
        accounts.append(closing[oc]['count'])
        credits += [np.nan] * 2
        states += [st] * 2
        branches += [branch] * 2
        dates += [date] * 2
        description_references += ['{} RQ DEP'.format(date)] * 2
    
    dates.append(date)

    description_references.append('{} RQ DEP'.format(date))
    branches.append('000')
    states.append('00')
    accounts.append('99998')
    debits.append(np.nan)
    credits.append(np.nan)
    
    report =  pd.DataFrame({'Date': dates,
                            'Type': ['G/L Account'] * len(debits),
                            'Account': accounts,
                            'St': states,
                            'Branch': branches,
                            'Dept': ['00'] * len(debits),
                            'Description Reference': description_references,
                            'Debits': debits,
                            'Credits': credits})
    
    report = report[~report.Account.str.startswith('?')]
    report = report[report.Debits != 0]
    return report

def clean_data(filename):
    if filename.endswith('csv'):
        df = pd.read_csv(filename, converters={'AcctCode':str, 'TitleCoNum':str})
    else:
        df = pd.read_excel(filename, converters={'AcctCode':str, 'TitleCoNum':str})
        
    
    df['Branch'] = df.TitleCoNum.map(lambda x: branches[x] if x in branches.keys() else '000')
    df['St'] = df['File Number'].apply(get_state)

    df.AcctCode.replace('40003', '40000', inplace=True)
    df.OrderCategory.replace(25, 8, inplace=True)
    df.dropna(subset=['Invoice Line Total'], inplace=True)

    df.loc[df[(df.AcctCode == '40000') & (df.OrderCategory.isin([2,5]))].index.tolist(), 'AcctCode'] = '40002'

    df['Dept'] = df.AcctCode.apply(get_dept)
    df['Date'] = pd.to_datetime(df['PaymentDate']).dt.strftime('%m/%d/%Y')

    return df[df['Invoice Line Total'] != 0]

def get_filename(df):
    os.makedirs('cash_receipts/', exist_ok=True)
    return 'cash_receipts/{} cash_receipts.xlsx'.format(df['Date'][0].replace('/','_'))

def fix_accounts(escrow, df):
    """Fix Account numbers for specific companies before creating sheet."""
    if escrow == 146:
        df.loc[df[df.Account == '96021'].index.tolist(), 'Account'] = '96024'
    elif escrow == 219:
        df.loc[df[df.Account == '96023'].index.tolist(), 'Account'] = '96020'
        df.loc[df[df.Account == '43502'].index.tolist(), 'Account'] = '43501'
    return df

def report_data(escrow, df):
    frame1 = report(escrow, df)
    frame2 = totals(escrow, df)
    journal = pd.concat([frame1, frame2], ignore_index=True)
    journal = fix_accounts(escrow, journal)
    s = journal[journal.Type == 'Bank Account'].index.values[0] + 3
    e = journal[journal.Account == '99998'].index.values[0] + 1
    f = '=SUM(I{}:I{})'.format(s,e)
    sheet = accounts[escrow]['sheet']
    return journal, accounts[escrow]['sheet'], f, e

def create_spreadsheet(filename, arr):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        print('Creating', filename.split('/')[-1] + '...')
        for i in range(len(arr)):
            print('Adding sheet', arr[i][1] + '...')
            arr[i][0].to_excel(writer, sheet_name=arr[i][1], index=False)
            workbook = writer.book
            worksheet = writer.sheets[arr[i][1]]
            num_format = workbook.add_format({'num_format':'##0.00'})
            worksheet.set_column(0, 2, 12)
            worksheet.set_column(3, 3, 3)
            worksheet.set_column(4, 5, 7)
            worksheet.set_column(6, 6, 15)
            worksheet.set_column(7, 7, 21)
            worksheet.set_column('I:J', 9, cell_format=num_format)
            worksheet.write_formula('J{}'.format(arr[i][-1] + 1), arr[i][-2])
        print('{} finished'.format(filename.split('/')[-1]))

def check_sheet(cash='sheets/cash.xls', fees='sheets/fee_master.xlsx', master=True):
    """
    Scans the worksheet and returns errors if any. if master is set to True it accounts that each file number in the master
    worksheet has been accounted for and the proper amount has been recorded.
    """
    errors, revisions = [],[]
    fees = pd.read_excel(fees)
    cash = pd.read_excel(cash)

    fee_files = set(fees['File'])

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

def check_balances(filename):
    xl = pd.ExcelFile(filename)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        rev_accts = [closing[x]['revenue'] for x in closing.keys()]
        x = round(df[df.Account.isin(rev_accts)]['Debits'].sum(), 2)
        y = round(df[df.Type == 'Bank Account']['Debits'].sum(), 2)
        if not x == y:
            print(sheet, 'Revenue Accounts:', x, 'Bank Account:', y)

def generate_journal(filename='sheets/cash.xls'):
    errors, revisions = check_sheet(master=True)
    if errors:
        for error in errors:
            print(error)
        return None
    for revision in revisions:
        print(revision)

    df = clean_data(filename)
    filename = get_filename(df)

    arr = list()
    for escrow, frame in df.groupby('EscrowBank'):
        arr.append(report_data(escrow, frame))

    arr = sorted(arr, key=lambda x: x[1])
    
    create_spreadsheet(filename, arr)

    check_balances(filename)

if __name__ == '__main__':
    branches = load_pickle('data/branches.pickle')
    closing = load_pickle('data/closing.pickle')
    accounts = load_pickle('data/accounts.pickle')
    generate_journal()
