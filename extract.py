import pandas as pd

file_name = 'sept.xlsx'


def load_data():
    data = file_name
    df = pd.read_excel(data)
    return df

def remove(df):
    print("Started Removing columns...")
    df.drop(df.columns[[7]], axis = 1, inplace = True)
    df.drop(['Purpose (Service Type)'], axis = 1, inplace = True)
    df.drop(df.loc[:, 'Funding Bank': 'Transmission Ref.'].columns, axis = 1, inplace = True)

    print("Finish removing columns...")

def rename_column(df):
    newdf = df.rename(columns={'ACCOUNT NUMBER ON THE BILL': 'Account No'})
    newdf.to_excel(file_name)

df = load_data()

def pre_processing():
    print("Started Pre-processing phase!!!")
    remove(df)
    rename_column(df)
    print("Finished Pre-processing phase!!!")

pre_processing()
