import pandas as pd
from send_email import send_mail

def import_excel(filepath):
    excel = pd.read_excel(filepath)
    
    for _, row in excel.iterrows():
        row = row.to_list()
        send_mail(row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8])
    
def check_cols(filepath):
    df = pd.read_excel(filepath)
    df = list(df.columns)
    df2 = pd.read_csv('col_check/cols.csv')
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




# check_cols("uploads/1eee3215-31fb-47fd-a2d9-cee7f883791bcols.xlsx")
# import_excel('sampledata copy.xlsx')