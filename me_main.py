# python 3.7
# %%
import sap_connection
import config as cf
import pyperclip
import time
import os
from datetime import date, timedelta, datetime
import dateutil
from calendar import monthrange
import pandas as pd
from sqlalchemy import create_engine
import pandas as pd

# %%

sap = None

engine = create_engine('mssql+pyodbc:'+cf.userandpass +
                       cf.database_p+'?driver=SQL+Server')


target_date = date.today()
if target_date.day <= 8:
    target_date = target_date - dateutil.relativedelta.relativedelta(months=1)

date_from = f'01.{target_date.month}.{target_date.year}'
date_to = f'{monthrange(target_date.year,target_date.month)[1]}.{target_date.month}.{target_date.year}'
datenow = date.today()
# %%


def login_p11():
    global sap
    global session
    sap = sap_connection.SAPConnection(
        user=cf.userP11, pwd=cf.pwdP11, system=cf.systemP11, language="EN")
    session = sap.session


def login_rpd():
    global sap
    global session
    sap = sap_connection.SAPConnection(
        user=cf.userRPD, pwd=cf.pwdRPD, system=cf.systemRPD, language="ES")
    session = sap.session


def login_p15():
    global sap
    global session
    sap = sap_connection.SAPConnection(
        user=cf.userP15, pwd=cf.pwdP15, system=cf.systemP15, language="EN")
    session = sap.session


def logout():
    try:
        global sap
        sap.break_connection()
        sap = None
    except:
        pass


def close_excel_by_force():
    import psutil

    for proc in psutil.process_iter():
        if proc.name() == "excel.exe" or proc.name() == "EXCEL.EXE":
            proc.kill()

def concat_files(all_files, outpath):
    li = []
    for filename in all_files:
        if os.path.exists(filename):
            df = pd.read_excel(filename, dtype=str)
            li.append(df)
    pd.concat(li, axis=0, ignore_index=True).to_excel(outpath, index=False)


def delete_files(all_files):
    import os
    for filename in all_files:
        if os.path.exists(filename):
            os.remove(filename)


def purchasing_documents_clipboard():
    df_ksb1 = pd.read_excel(
        cf.path_data_local+r'\RPD\ksb1_rpd.xlsx', converters={'Documento compras': str})
    df_kob1 = pd.read_excel(
        cf.path_data_local+r'\RPD\kob1_rpd.xlsx', converters={'Documento compras': str})
    doc_list = df_ksb1['Documento compras'].append(
        df_kob1['Documento compras'], ignore_index=True)
    doc_list = doc_list.dropna().drop_duplicates()
    pyperclip.copy('\r\n'.join(doc_list))


def excel_to_db(file_name, table_name, targetcolumns, source, if_exists='replace', all_text=False):
    datenow = date.today()
    if all_text:
        df = pd.read_excel(file_name, dtype=str)
    else:
        df = pd.read_excel(file_name)
    df['Source'] = source
    df['Load_date'] = f'{datenow.year}-{datenow.month:02d}-{datenow.day:02d}'
    df.columns = targetcolumns
    df.to_sql(table_name, schema='MandE', con=engine,
              if_exists=if_exists, index=False)


# %%
#get data ksb1 RPD
def download_ksb1_rpd():
    file_name = cf.path_data_local+r'\RPD\ksb1_rpd.xlsx'
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKSB1"

    session.findById("wnd[0]").sendVKey(0)
    try:
        session.findById(
            "wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "VT01"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass

    pyperclip.copy(cf.clase_coste_rpd)

    try:
        session.findById("wnd[0]/usr/btn%_KSTAR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
    except:
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nKSB1"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btn%_KSTAR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtKOSTL-LOW").text = "E710000000"
    session.findById("wnd[0]/usr/ctxtKOSTL-HIGH").text = "E889999999"
    session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = date_from
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = date_to
    session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/AUTOME"
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").caretPosition = 10
    session.findById("wnd[0]/usr/btnBUT1").press()
    session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "99999999"
    session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    #Copy data in clipboard
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    #Read data from clipboard y split lines by |
    df_ksb1 = pd.read_clipboard(sep="|", header =3, skipfooter=1, usecols= list(range(1,39)), engine='python', dtype=str, quoting=3 )

    df_ksb1 = df_ksb1.drop(0)
    df_ksb1 = df_ksb1.apply(lambda x : x.str.strip() if x.dtype == "object" else x)

    df_ksb1.columns = cf.COLUMNS_RPD_ME

    df_ksb1['Source'] = 'ksb1_rpd'
    df_ksb1['Load_date'] = f'{datenow.year}-{datenow.month:02d}-{datenow.day:02d}'
    #Insert data
    df_ksb1.to_sql('EXT_rpd', schema='MandE', con=engine, if_exists='append', index=False)


#get data kob1 RPD
def download_kob1_rpd():
    file_name = cf.path_data_local+r'\RPD\kob1_rpd.txt'
    datapath = cf.path_data_local_rpd
    name = 'kob1_rpd.txt'

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKOB1"
    session.findById("wnd[0]").sendVKey(0)
    try:
        session.findById(
            "wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "VT01"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass

    # Get Clase Costo into clipboard
    pyperclip.copy(cf.clase_coste_rpd)
    try:
        session.findById("wnd[0]/usr/btn%_KSTAR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
    except:
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nKOB1"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btn%_KSTAR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

    session.findById("wnd[0]/usr/ctxtAUFNR-LOW").text = "1"
    session.findById("wnd[0]/usr/ctxtAUFNR-HIGH").text = "999999999999"
    session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = date_from
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = date_to
    session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/AUTOME"
    session.findById("wnd[0]/usr/ctxtP_DISVAR").setFocus()
    session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 7
    session.findById("wnd[0]/usr/btnBUT1").press()
    session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "99999999"
    session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    #Extract file
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = datapath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = name
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    #Read file
    with open(file_name) as f:
        contents = f.read()

    #Replace problem strings
    for key, value in cf.replaceVal.items():
        if value[0] in contents:
            contents = contents.replace(value[0],value[1])

    #Delete the first no important lines 
    contents = contents[569:]

    #Split each line by '\n' and store it in a list position 
    contentslist = contents.split('\n')

    #The first 4 positions are no important
    del contentslist[0:4]

    i = 2
    df_kob1 =  pd.DataFrame(columns=cf.COLUMNS_RPD_ME)

    #pd.set_option('display.max_columns', None)
    #pd.set_option('display.max_rows', None)

    #Split each line by | and store it like a new row in the dataset
    while i < len(contentslist): #5:
        line = contentslist[i]
        LineList = line.split('|')
        
        LineList = {'FirstCol':LineList[0], 'Sociedad':LineList[1], 'Centro':LineList[2], 'Segmento':LineList[3], 'Objeto_del_interlocutor':LineList[4], 'Denom_del_objeto_del_inter':LineList[5], 'Clase_de_coste':LineList[6], 'Denom_clase_de_coste':LineList[7], 'Descrip_clases_coste':LineList[8], 'Clase_cta_de_contrapartida':LineList[9], 'Cta_contrapartida':LineList[10], 'Denom_cuenta_contrapartida':LineList[11], 'Denom_cuenta_contrapartida1':LineList[12], 'Clase_de_documento':LineList[13], 'Documento_compras':LineList[14], 'Fe_contabilizacion':LineList[15], 'Ejercicio':LineList[16],'Periodo':LineList[17], 'Material':LineList[18], 'Texto_breve_de_material':LineList[19], 'Texto_de_pedido':LineList[20], 'Denominacion':LineList[21], 'Ud_cantidad_contab':LineList[22], 'Cantidad_total_reg':LineList[23], 'Moneda_del_objeto':LineList[24], 'Valor_moneda_objeto':LineList[25], 'Moneda_transacción':LineList[26], 'Valor_mon_tr':LineList[27], 'Moneda_sociedad_CO':LineList[28], 'Valor_mon_soc_CO':LineList[29], 'Objeto':LineList[30], 'Denom_del_objeto':LineList[31], 'Texto_de_cabecera_de_doc':LineList[32], 'Fecha_de_doc':LineList[33], 'Numero_de_documento':LineList[34], 'Area_funcional':LineList[35], 'Tipo_de_valor':LineList[36], 'Orden':LineList[37], 'Posicion':LineList[38], 'LastCol':LineList[39]}
        df_kob1 = df_kob1.append(LineList,  ignore_index=True)
        i = i+1
    
    df_kob1['Source'] = 'kob1_rpd'
    df_kob1['Load_date'] = f'{datenow.year}-{datenow.month:02d}-{datenow.day:02d}'

    #Insert in data in table: EXT_rpd
    df_kob1.to_sql('EXT_rpd', schema='MandE', con=engine, if_exists='replace', index=False)


def download_mara_rpd():
    file_name = cf.path_data_local+r'\RPD\mara_rpd.xlsx'
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16N"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "MARA"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
    session.findById("wnd[0]/tbar[1]/btn[18]").press()
    session.findById(
        "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,1]").selected = True
    session.findById(
        "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,11]").selected = True
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById(
        "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    #Copy data in clipboard
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    #Read data from clipboard y split lines by |
    df_mara = pd.read_clipboard(sep="|", header =3, skipfooter=1, usecols= list(range(1,3)), engine='python', dtype=str)
    
    df_mara = df_mara.drop(0)
    df_mara = df_mara.apply(lambda x : x.str.strip() if x.dtype == "object" else x)

    df_mara['Source'] = 'mara_rpd'
    df_mara['Load_date'] = f'{datenow.year}-{datenow.month:02d}-{datenow.day:02d}'

    #Rename columns
    df_mara.columns = cf.COLUMNS_RPD_MARA
    #Insert data
    df_mara.to_sql('EXT_rpd_mara', schema='MandE', con=engine, if_exists='replace', index=False)
    


def download_t023t_rpd():
    file_name = cf.path_data_local+r'\RPD\t023t_rpd.xlsx'
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16N"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "T023T"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
    session.findById(
        "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").text = "ES"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    #Export data
    session.findById(
        "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    #Read data from clipboard y split lines by |
    df_t023t = pd.read_clipboard(sep="|", header =3, skipfooter=1, usecols= list(range(1,5)), engine='python', dtype=str)
    
    df_t023t = df_t023t.drop(0)
    df_t023t = df_t023t.apply(lambda x : x.str.strip() if x.dtype == "object" else x)

    df_t023t['Source'] = 't023t_rpd'
    df_t023t['Load_date'] = f'{datenow.year}-{datenow.month:02d}-{datenow.day:02d}'

    #Rename columns
    df_t023t.columns = cf.COLUMNS_RPD_T023T

    #Insert data
    df_t023t.to_sql('EXT_rpd_mb52', schema='MandE', con=engine, if_exists='replace', index=False)
    


def download_me2n_rpd():
    file_name = cf.path_data_local+r'\RPD\me2n_rpd.xlsx'
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nme2n"
    session.findById("wnd[0]").sendVKey(0)
    # Read Documentos de compra
    purchasing_documents_clipboard()
    session.findById("wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    #Copy data in clipboard
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    #Read data from clipboard y split lines by |
    df_me2n = pd.read_clipboard(sep="|", header =3, skipfooter=1, usecols= list(range(1,21)), engine='python', dtype=str)
    
    df_me2n = df_me2n.drop(0)
    df_me2n = df_me2n.apply(lambda x : x.str.strip() if x.dtype == "object" else x)

    df_me2n['Source'] = 'me2n_rpd'
    df_me2n['Load_date'] = f'{datenow.year}-{datenow.month:02d}-{datenow.day:02d}'
    #Rename columns
    df_me2n.columns = cf.COLUMNS_RPD_ME2N
    #Insert data 
    df_me2n.to_sql('EXT_rpd_me2n', schema='MandE', con=engine, if_exists='replace', index=False)


# %%
def download_ksb1_p15():
    file_name = cf.path_data_local+r'\ksb1_p15.txt'
    fromcostcenter = '4000000'
    tocostcenter = '6999999'
    costelementgroup = 'H73400'
    layout = '/TSCAMMEV1.0'
    datapath = cf.path_data_local
    name = 'ksb1_p15.txt'
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKSB1"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtKOSTL-LOW").text = fromcostcenter
    session.findById("wnd[0]/usr/ctxtKOSTL-HIGH").text = tocostcenter
    session.findById("wnd[0]/usr/ctxtKOAGR").text = costelementgroup
    session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = date_from
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = date_to
    session.findById("wnd[0]/usr/ctxtP_DISVAR").text = layout
    session.findById("wnd[0]/usr/btnBUT1").press()
    session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "99999999"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    #Extract txt file separated by |
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = datapath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = name
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    #Read txt file
    with open(file_name) as f:
        contents = f.read()

    #Replace wrong strings
    for key, value in cf.replaceValP15.items():
        if value[0] in contents:
            contents = contents.replace(value[0],value[1])

    #Delete not useful characteres
    contents = contents[618:]

    #split each line by '\n' and store it in a list position 
    contentslist = contents.split('\n')

    #Delete column names row
    del contentslist[0:1]
    i=0
    df_ksb1 =  pd.DataFrame(columns=cf.COLUMNS_P15_KSB1)

    #Split each line by | and store it like a new row in the dataset
    while i < len(contentslist)-1: #5:
        line = contentslist[i]
        LineList = line.split('|')

        LineList = {'Company_code':LineList[1], 'Plant':LineList[2], 'Cost_center':LineList[3], 'CO_object_name':LineList[4], 'Cost_element':LineList[5], 'Cost_element_name':LineList[6], 'Cost_element_descr':LineList[7], 'Offsetting_account_type':LineList[8], 'Offsetting_acct_no':LineList[9], 'Name_of_offsetting_account':LineList[10], 'Name_of_offsetting_account1':LineList[11], 'Document_type':LineList[12], 'Purchasing_document':LineList[13], 'Posting_date':LineList[14], 'Fiscal_year':LineList[15] ,'Period':LineList[16], 'Material':LineList[17], 'Material_description':LineList[18],'Purchase_order_text':LineList[19], 'Name':LineList[20], 'Posted_unit_of_meas':LineList[21], 'Total_quantity':LineList[22], 'Object_currency':LineList[23], 'Value_in_obj_crcy':LineList[24], 'Transaction_currency':LineList[25], 'Value_trancurr':LineList[26], 'Report_currency':LineList[27], 'Val_in_rep_cur':LineList[28], 'Cost_element_group':LineList[29], 'Cost_element_group_name':LineList[30], 'Document_header_text':LineList[31], 'Document_date':LineList[32], 'Document_number':LineList[33], 'Functional_area':LineList[34], 'Material_group':LineList[35], 'Material_group_desc':LineList[36], 'Value_type':LineList[37]}
        df_ksb1 = df_ksb1.append(LineList,  ignore_index=True)
        i = i+1


    df_ksb1['Source'] = 'P15_me_Ksb1'
    df_ksb1['Load_date'] = f'{datenow.year}-{datenow.month:02d}-{datenow.day:02d}'

    df_ksb1.to_sql('EXT_p15_ksb1', schema='MandE', con=engine, if_exists='replace', index=False)


def download_koc2_p15():
    file_name = r'\p15_koc2_me.xlsx'
    from_order = "100000000000"
    to_order = "200000000000"
    costelementgroup = 'H73400'
    layout = '/TSCSAPPM'
    datapath = cf.path_data_local
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKOC2"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/lbl[5,6]").setFocus()
    session.findById("wnd[0]/usr/lbl[5,6]").caretPosition = 0
    session.findById("wnd[0]").sendVKey(2)
    session.findById("wnd[0]/usr/lbl[12,8]").setFocus()
    session.findById("wnd[0]/usr/lbl[12,8]").caretPosition = 0
    session.findById("wnd[0]").sendVKey(2)
    session.findById("wnd[0]/usr/ctxtAUFNR-LOW").text = from_order
    session.findById("wnd[0]/usr/ctxtAUFNR-HIGH").text = to_order
    session.findById("wnd[0]/usr/ctxtKOAGR").text = costelementgroup
    session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = date_from
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = date_to
    session.findById("wnd[0]/usr/ctxtP_DISVAR").text = layout
    session.findById("wnd[0]/usr/btnBUT1").press()
    session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "99999999"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    #Export data
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/usr/radRB_OTHERS").setFocus()
    session.findById("wnd[1]/usr/radRB_OTHERS").select()
    session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
    session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "10"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = datapath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    time.sleep(30)
    close_excel_by_force()

    #From excel to DB
    excel_to_db(datapath+file_name, 'EXT_p15_koc2_me',
                cf.COLUMNS_P15_KOC2, 'P15_me_Koc2', if_exists='replace', all_text=True)

# %%


def download_ksb1_p11():
    delete_files(cf.p11_me_files)
    fromcostcenter = 'CZ00000'
    tocostcenter = 'PLS9999'
    costelementgrouplist = ['MEREG', 'I35300', 'I35600']
    layout = '/TSCEUMEV1.0'
    datapath = cf.path_data_local
    outsufix = "p11_ksb1_me_"
    for costelementgroup in costelementgrouplist:
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nKSB1"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtKOSTL-LOW").text = fromcostcenter
            session.findById("wnd[0]/usr/ctxtKOSTL-HIGH").text = tocostcenter
            session.findById("wnd[0]/usr/ctxtKOAGR").text = costelementgroup
            session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = date_from
            session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = date_to
            session.findById("wnd[0]/usr/ctxtP_DISVAR").text = layout
            session.findById("wnd[0]/usr/btnBUT1").press()
            session.findById(
                "wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "99999999"
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
            session.findById("wnd[1]/usr/radRB_OTHERS").setFocus()
            session.findById("wnd[1]/usr/radRB_OTHERS").select()
            session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
            session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "31"
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = datapath
            session.findById(
                "wnd[1]/usr/ctxtDY_FILENAME").text = outsufix+costelementgroup+".xlsx"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except:
            print("Unable to download P11"+costelementgroup)
    time.sleep(15)
    close_excel_by_force()
    concat_files(cf.p11_me_files, cf.path_data_local+r'\p11_ksb1_me.xlsx')
    #From excel to DB
    excel_to_db(cf.path_data_local+r'\p11_ksb1_me.xlsx',
                'EXT_p11_ksb1', cf.COLUMNS_P11_KSB1, 'P11_me_Ksb1', if_exists='replace', all_text=True)

# %%

def main():
    try:
        print(date_from)
        print(date_to)
        logout()
        login_p11()
        download_ksb1_p11()
        logout()
        login_rpd()
        download_mara_rpd()
        download_t023t_rpd()
        download_me2n_rpd()
        download_kob1_rpd()
        download_ksb1_rpd()
        logout()
        login_p15()
        download_ksb1_p15()
        logout()
        login_p15()
        download_koc2_p15()
        logout()
    except Exception as ex:
        print(ex)
        logout()


# %%
if __name__ == '__main__':
    main()