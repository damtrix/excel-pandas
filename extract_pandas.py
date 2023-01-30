from openpyxl import load_workbook, workbook
import pandas as pd

file_name = 'new/December.xlsx'


def load_data():
    data = file_name
    df = pd.read_excel(data)
    return df


def remove():
    print("Started Removing columns...")
    book = load_workbook(file_name)
    sheet1 = book['Blank A4 Landscape']
    sheet1.delete_cols(7)
    sheet1.delete_cols(7)
    sheet1.delete_cols(11, 10)
    book.save(file_name)

    
    print("Finish removing columns...")

def rename_column(df):
    newdf = df.rename(columns={'ACCOUNT NUMBER ON THE BILL': 'Account No'})
    #newdf.drop(newdf.columns[[0]], axis = 1, inplace = True)
    newdf.to_excel(file_name, index = None)

def pre_processing():
    print("Started Pre-processing phase!!!")
    remove()
    df = load_data()
    rename_column(df)
    print("Finished Pre-processing phase!!!")

pre_processing()

df = load_data()

def sum_net_amount(district, data):
    net_total = round((data['Net Amount'].sum()), 2)
    total.loc[district] = [net_total]


filter_asokoro = df.loc[(df['Account No'].str.startswith('A04', na = False)) | (df['Account No'].str.startswith('Ao4')) | (df['Account No'].str.startswith('ao4')) | (df['Account No'].str.startswith('a04')) | (df['Account No'].str.startswith('aO4')) | (df['Account No'].str.startswith('AO4')) | (df['Account No'].str.startswith('AO0'))]

# col_list = ['Amount', 'Total Fee', 'Net Amount']

# sum = filter_asokoro.sum(col_list)
# sum.name = 'Sum'
# filter_asokoro = filter_asokoro.append(sum.transpose())

# df['sum'] = filter_asokoro[col_list].sum(axis = 1)



filter_cbd = df.loc[(df['Account No'].str.startswith('A00', na = False)) | (df['Account No'].str.startswith('a00')) | (df['Account No'].str.startswith('A0o')) | (df['Account No'].str.startswith('Ao0')) | (df['Account No'].str.startswith('a0o')) | (df['Account No'].str.startswith('ao0')) | (df['Account No'].str.startswith('aoo')) | (df['Account No'].str.startswith('Aoo')) | (df['Account No'].str.startswith('AOO')) | (df['Account No'].str.startswith('aOO'))]



filter_garki_1 = df.loc[(df['Account No'].str.startswith('A01', na = False)) | (df['Account No'].str.startswith('a01')) | (df['Account No'].str.startswith('Ao1')) | (df['Account No'].str.startswith('ao1')) | (df['Account No'].str.startswith('AO1')) | (df['Account No'].str.startswith('aO1')) | (df['Account No'].str.startswith('A0i')) | (df['Account No'].str.startswith('A0I')) | (df['Account No'].str.startswith('a0i')) | (df['Account No'].str.startswith('a0I')) | (df['Account No'].str.startswith('Aoi')) | (df['Account No'].str.startswith('AoI'))]

# garki_1_net_sum = round((filter_garki_1['Net Amount'].sum()), 2)
# total.loc['CBD'] = [garki_1_net_sum]

filter_garki_2 = df.loc[(df['Account No'].str.startswith('A03', na = False)) | (df['Account No'].str.startswith('a03')) | (df['Account No'].str.startswith('AO3')) | (df['Account No'].str.startswith('aO3')) | (df['Account No'].str.startswith('Ao3')) | (df['Account No'].str.startswith('ao3'))]

filter_wuse_1 = df.loc[(df['Account No'].str.startswith('A02', na = False)) | (df['Account No'].str.startswith('a02')) | (df['Account No'].str.startswith('AO2')) | (df['Account No'].str.startswith('aO2')) | (df['Account No'].str.startswith('Ao2')) | (df['Account No'].str.startswith('ao2'))]

filter_wuse_2 = df.loc[(df['Account No'].str.startswith('WII', na = False)) | (df['Account No'].str.startswith('W11')) | (df['Account No'].str.startswith('wii')) | (df['Account No'].str.startswith('w11')) | (df['Account No'].str.startswith('w1i')) | (df['Account No'].str.startswith('wi1')) | (df['Account No'].str.startswith('W1i')) | (df['Account No'].str.startswith('Wi1')) | (df['Account No'].str.startswith('W11')) | (df['Account No'].str.startswith('Wii')) | (df['Account No'].str.startswith('Wll')) | (df['Account No'].str.startswith('WIi')) | (df['Account No'].str.startswith('W2')) | (df['Account No'].str.startswith('WI1'))]

filter_maitama = df.loc[(df['Account No'].str.startswith('MAI', na = False)) | (df['Account No'].str.startswith('mai')) | (df['Account No'].str.startswith('Mai')) | (df['Account No'].str.startswith('Mai')) | (df['Account No'].str.startswith('mAi')) | (df['Account No'].str.startswith('MAi')) | (df['Account No'].str.startswith('mAI')) | (df['Account No'].str.startswith('maI')) | (df['Account No'].str.startswith('MaI')) | (df['Account No'].str.startswith('MA1')) | (df['Account No'].str.startswith('Ma1'))]


filter_jabi = df.loc[(df['Account No'].str.startswith('B04', na = False)) | (df['Account No'].str.startswith('b04')) | (df['Account No'].str.startswith('Bo4')) | (df['Account No'].str.startswith('BO4')) | (df['Account No'].str.startswith('bo4')) | (df['Account No'].str.startswith('bO4'))]

filter_gwarinpa = df.loc[(df['Account No'].str.startswith('C02', na = False)) | (df['Account No'].str.startswith('Co2')) | (df['Account No'].str.startswith('CO2')) | (df['Account No'].str.startswith('c02')) | (df['Account No'].str.startswith('cO2')) | (df['Account No'].str.startswith('co2'))]

filter_jahi = df.loc[(df['Account No'].str.startswith('B08', na = False)) | (df['Account No'].str.startswith('BO8')) | (df['Account No'].str.startswith('Bo8')) | (df['Account No'].str.startswith('b08')) | (df['Account No'].str.startswith('bo8')) | (df['Account No'].str.startswith('bO8'))]

filter_kado = df.loc[(df['Account No'].str.startswith('B09', na = False)) | (df['Account No'].str.startswith('BO9')) | (df['Account No'].str.startswith('Bo9')) | (df['Account No'].str.startswith('b09')) | (df['Account No'].str.startswith('bo9')) | (df['Account No'].str.startswith('bO9'))]

filter_utako = df.loc[(df['Account No'].str.startswith('B05', na = False)) | (df['Account No'].str.startswith('BO5')) | (df['Account No'].str.startswith('Bo5')) | (df['Account No'].str.startswith('b05')) | (df['Account No'].str.startswith('bo5')) | (df['Account No'].str.startswith('bO5'))]

filter_gaduwa = df.loc[(df['Account No'].str.startswith('B13', na = False)) | (df['Account No'].str.startswith('BI3')) | (df['Account No'].str.startswith('b13')) | (df['Account No'].str.startswith('bI3')) | (df['Account No'].str.startswith('bi3')) | (df['Account No'].str.startswith('Bi3'))]

filter_durumi = df.loc[(df['Account No'].str.startswith('B02', na = False)) | (df['Account No'].str.startswith('BO2')) | (df['Account No'].str.startswith('Bo2')) | (df['Account No'].str.startswith('b05')) | (df['Account No'].str.startswith('bo5')) | (df['Account No'].str.startswith('bO5'))]

filter_gudu_apo = df.loc[(df['Account No'].str.startswith('B01', na = False)) | (df['Account No'].str.startswith('BO1')) | (df['Account No'].str.startswith('B0I')) | (df['Account No'].str.startswith('b01')) | (df['Account No'].str.startswith('bO1')) | (df['Account No'].str.startswith('b0I')) | (df['Account No'].str.startswith('bOI')) | (df['Account No'].str.startswith('b0i')) | (df['Account No'].str.startswith('boi')) | (df['Account No'].str.startswith('bOi')) | (df['Account No'].str.startswith('bo1'))]

filter_mabushi = df.loc[(df['Account No'].str.startswith('B06', na = False)) | (df['Account No'].str.startswith('BO6')) | (df['Account No'].str.startswith('Bo6')) | (df['Account No'].str.startswith('b06')) | (df['Account No'].str.startswith('bO6')) | (df['Account No'].str.startswith('bo6'))]

filter_life_camp = df.loc[(df['Account No'].str.startswith('L00', na = False)) | (df['Account No'].str.startswith('LOO')) | (df['Account No'].str.startswith('L0O')) | (df['Account No'].str.startswith('LO0')) | (df['Account No'].str.startswith('loo')) | (df['Account No'].str.startswith('l00')) | (df['Account No'].str.startswith('lO0')) | (df['Account No'].str.startswith('lOO'))]

filter_wuye = df.loc[(df['Account No'].str.startswith('B03', na = False)) | (df['Account No'].str.startswith('BO3')) | (df['Account No'].str.startswith('Bo3')) | (df['Account No'].str.startswith('b03')) | (df['Account No'].str.startswith('b03')) | (df['Account No'].str.startswith('bo3')) | (df['Account No'].str.startswith('bo3'))]

filter_katampe_ext = df.loc[(df['Account No'].str.startswith('B19', na = False)) | (df['Account No'].str.startswith('BI9')) | (df['Account No'].str.startswith('b19')) | (df['Account No'].str.startswith('bi9')) | (df['Account No'].str.startswith('Bi9'))]

filter_lugbe = df.loc[(df['Account No'].str.startswith('LUG', na = False)) | (df['Account No'].str.startswith('lug')) | (df['Account No'].str.startswith('Lug'))]

filter_dutse = df.loc[(df['Account No'].str.startswith('B14', na = False)) | (df['Account No'].str.startswith('BI4')) | (df['Account No'].str.startswith('b14')) | (df['Account No'].str.startswith('bi4')) | (df['Account No'].str.startswith('bI4'))]

filter_exp = df.loc[(df['Account No'].str.startswith('EXP', na = False)) | (df['Account No'].str.startswith('Exp')) | (df['Account No'].str.startswith('exp'))]

filter_guzape = df.loc[(df['Account No'].str.startswith('A09', na = False)) | (df['Account No'].str.startswith('AO9')) | (df['Account No'].str.startswith('a09')) | (df['Account No'].str.startswith('ao9')) | (df['Account No'].str.startswith('aO9'))]

filter_kaura = df.loc[(df['Account No'].str.startswith('B11', na = False)) | (df['Account No'].str.startswith('BII')) | (df['Account No'].str.startswith('b11')) | (df['Account No'].str.startswith('Bii')) | (df['Account No'].str.startswith('bii')) | (df['Account No'].str.startswith('b1i')) | (df['Account No'].str.startswith('bi1'))]

filter_airport = df.loc[(df['Account No'].str.startswith('AIR', na = False)) | (df['Account No'].str.startswith('air')) | (df['Account No'].str.startswith('Air')) | (df['Account No'].str.startswith('a1r'))]

filter_idu = df.loc[(df['Account No'].str.startswith('IDU', na =False)) | (df['Account No'].str.startswith('idu')) | (df['Account No'].str.startswith('Idu')) | (df['Account No'].str.startswith('iDu'))]

filter_emb =  df.loc[(df['Account No'].str.startswith('emb', na = False)) | (df['Account No'].str.startswith('EMB')) | (df['Account No'].str.startswith('Emb')) | (df['Account No'].str.startswith('emB')) | (df['Account No'].str.startswith('eMb'))]

filter_lokogoma = df.loc[((df['Account No'].str.startswith('C09', na = False))) | (df['Account No'].str.startswith('c09')) | (df['Account No'].str.startswith('CO9')) | (df['Account No'].str.startswith('cO9')) | (df['Account No'].str.startswith('co9')) | (df['Account No'].str.startswith('Co9'))]

filter_dawaki = df.loc[((df['Account No'].str.startswith('DAW', na = False))) | ((df['Account No'].str.startswith('Daw'))) | ((df['Account No'].str.startswith('daw')))]

filter_katampe_main = df.loc[((df['Account No'].str.startswith('B07', na = False))) | ((df['Account No'].str.startswith('b07'))) | ((df['Account No'].str.startswith('BO7'))) | ((df['Account No'].str.startswith('bO7')))]

filter_kubwa = df.loc[((df['Account No'].str.startswith('KW1', na = False))) |  ((df['Account No'].str.startswith('KWI'))) | ((df['Account No'].str.startswith('kw1'))) | ((df['Account No'].str.startswith('Kw1', na = False)))]

filter_account_name = df.loc[((df['Account No'].str.startswith('ACCOUNT', na = False)))]

filter_serial = df.loc[((df['S/No'].str.startswith('S', na = False)))]

filter_grand = df.loc[((df['Reference No. (RRR)'].str.startswith('GRAND TOTA', na = False)))]

df.dropna(subset=['Reference No. (RRR)'], inplace = True)

filter_no_name = df.loc[(~df['Account No'].isin(filter_cbd['Account No'])) & (~df['Account No'].isin(filter_garki_1['Account No'])) & (~df['Account No'].isin(filter_garki_2['Account No'])) & (~df['Account No'].isin(filter_wuse_1['Account No'])) & (~df['Account No'].isin(filter_wuse_2['Account No'])) & (~df['Account No'].isin(filter_maitama['Account No'])) & (~df['Account No'].isin(filter_jabi['Account No'])) & (~df['Account No'].isin(filter_gwarinpa['Account No'])) & (~df['Account No'].isin(filter_jahi['Account No'])) & (~df['Account No'].isin(filter_kado['Account No'])) & (~df['Account No'].isin(filter_utako['Account No'])) & (~df['Account No'].isin(filter_gaduwa['Account No'])) & (~df['Account No'].isin(filter_durumi['Account No'])) & (~df['Account No'].isin(filter_gudu_apo['Account No'])) & (~df['Account No'].isin(filter_mabushi['Account No'])) & (~df['Account No'].isin(filter_life_camp['Account No'])) & (~df['Account No'].isin(filter_wuye['Account No'])) & (~df['Account No'].isin(filter_katampe_ext['Account No'])) & (~df['Account No'].isin(filter_lugbe['Account No'])) & (~df['Account No'].isin(filter_dutse['Account No'])) & (~df['Account No'].isin(filter_exp['Account No'])) & (~df['Account No'].isin(filter_guzape['Account No'])) & (~df['Account No'].isin(filter_kaura['Account No'])) & (~df['Account No'].isin(filter_airport['Account No'])) & (~df['Account No'].isin(filter_idu['Account No'])) & (~df['Account No'].isin(filter_emb['Account No'])) & (~df['Account No'].isin(filter_lokogoma['Account No'])) & (~df['Account No'].isin(filter_asokoro['Account No'])) & (~df['Account No'].isin(filter_dawaki['Account No'])) & (~df['Account No'].isin(filter_account_name['Account No'])) & (~df['Account No'].isin(filter_kubwa['Account No'])) & (~df['S/No'].isin(filter_serial['S/No'])) & (~df['Reference No. (RRR)'].isin(filter_grand['Reference No. (RRR)'])) & (~df['Account No'].isin(filter_katampe_main['Account No']))]

no_name_net_sum = filter_no_name['Net Amount'].sum()

total = pd.DataFrame({"Amount": [round(no_name_net_sum, 2)]}, index=['NO NAME'])

sum_net_amount('KAURA', filter_kaura)
sum_net_amount('GADUWA', filter_gaduwa)
sum_net_amount('DURUMI', filter_durumi)
sum_net_amount('GUDU APO', filter_gudu_apo)
sum_net_amount('GWARINPA', filter_gwarinpa)
sum_net_amount('MABUSHI', filter_mabushi)
sum_net_amount('LIFE CAMP', filter_life_camp)
sum_net_amount('WUYE', filter_wuye)
sum_net_amount('KADO', filter_kado)
sum_net_amount('ASOKORO', filter_asokoro)
sum_net_amount('CBD', filter_cbd)
sum_net_amount('GARKI I', filter_garki_1)
sum_net_amount('GARKI II', filter_garki_2)
sum_net_amount('WUSE I', filter_wuse_1)
sum_net_amount('WUSE II', filter_wuse_2)
sum_net_amount('MAITAMA', filter_maitama)
sum_net_amount('JABI', filter_jabi)
sum_net_amount('UTAKO', filter_utako)
sum_net_amount('EXP', filter_exp)
sum_net_amount('LOKOGOMA', filter_lokogoma)
sum_net_amount('KATAMPE EXT', filter_katampe_ext)
sum_net_amount('JAHI', filter_jahi)
sum_net_amount('LUGBE', filter_lugbe)
sum_net_amount('DAWAKI', filter_dawaki)
sum_net_amount('DUTSE', filter_dutse)
sum_net_amount('GUZAPE', filter_guzape)
sum_net_amount('IDU', filter_idu)
sum_net_amount('KATAMPE MAIN', filter_katampe_main)
sum_net_amount('AIRPORT', filter_airport)
sum_net_amount('EMBASSY', filter_emb)
sum_net_amount('KUBWA', filter_kubwa)




with pd.ExcelWriter(file_name) as writer:
    df.to_excel(writer, sheet_name='sheet 1', index = None)
    filter_asokoro.to_excel(writer, sheet_name='ASOKORO', index = None)
    filter_cbd.to_excel(writer, sheet_name='CBD', index = None)
    filter_garki_1.to_excel(writer, sheet_name='GARKI I', index = None)
    filter_garki_2.to_excel(writer, sheet_name='GARKI II', index = None)
    filter_wuse_1.to_excel(writer, sheet_name='WUSE I', index = None)
    filter_wuse_2.to_excel(writer, sheet_name='WUSE II', index = None)
    filter_maitama.to_excel(writer, sheet_name='MAITAMA', index = None)
    filter_jabi.to_excel(writer, sheet_name='JABI', index = None)
    filter_gwarinpa.to_excel(writer, sheet_name='GWARINPA', index = None)
    filter_jahi.to_excel(writer, sheet_name='JAHI', index = None)
    filter_kado.to_excel(writer, sheet_name='KADO', index = None)
    filter_utako.to_excel(writer, sheet_name='UTAKO', index = None)
    filter_gaduwa.to_excel(writer, sheet_name='GADUWA', index = None)
    filter_durumi.to_excel(writer, sheet_name='DURUMI', index = None)
    filter_gudu_apo.to_excel(writer, sheet_name='GUDU APO', index = None)
    filter_mabushi.to_excel(writer, sheet_name='MABUSHI', index = None)
    filter_life_camp.to_excel(writer, sheet_name='LIFE CAMP', index = None)
    filter_wuye.to_excel(writer, sheet_name='WUYE', index = None)
    filter_katampe_ext.to_excel(writer, sheet_name='KATAMPE EXT', index = None)
    filter_lugbe.to_excel(writer, sheet_name='LUGBE', index = None)
    filter_dutse.to_excel(writer, sheet_name='DUTSE', index = None)
    filter_exp.to_excel(writer, sheet_name='EXP', index = None)
    filter_guzape.to_excel(writer, sheet_name='GUZAPE', index = None)
    filter_kaura.to_excel(writer, sheet_name='KAURA', index = None)
    filter_airport.to_excel(writer, sheet_name='AIRPORT', index = None)
    filter_idu.to_excel(writer, sheet_name='IDU', index = None)
    filter_emb.to_excel(writer, sheet_name='EMBASSY', index = None)
    filter_lokogoma.to_excel(writer, sheet_name='LOKOGOMA', index = None)
    filter_dawaki.to_excel(writer, sheet_name='DAWAKI', index = None)
    filter_katampe_main.to_excel(writer, sheet_name='KATAMPE MAIN', index = None)
    filter_kubwa.to_excel(writer, sheet_name='KUBWA', index = None)
    filter_no_name.to_excel(writer, sheet_name='NO NAME', index = None)
    total.to_excel(writer, sheet_name='TOTAL')

