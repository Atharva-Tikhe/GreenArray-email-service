import pandas as pd
import os
from send_email import send_mail

def import_excel(filepath):
    excel = pd.read_excel(filepath)
    
    for _, row in excel.iterrows():
        row = row.to_list()
        row[9] = [bcc.strip() for bcc in row[9].split(',')]
        send_mail(row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9])
    
def check_cols(filepath):
    df = pd.read_excel(filepath)
    df = list(df.columns)
    df2 = pd.read_csv(os.path.join('col_check', 'cols.csv'))
    df2 = list(df2.columns)
    
    col_check = []

    for col in range(len(df)):
        df[col] = df[col].strip()
        df2[col] = df[col].strip()
        if df[col] == df2[col]:
            col_check.append(0)
        else:
            col_check.append(1)
    
    if sum(col_check) == 0:
        print(col_check)
        return True
    else:
        print(col_check)
        raise Exception(f"ColumnMismatchErrror: The columns do not match {col_check}")




# check_cols("sampledata copy.xlsx")
# import_excel('sampledata copy.xlsx')