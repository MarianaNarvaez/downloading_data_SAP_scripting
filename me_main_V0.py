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

# %%

sap = None

target_date = date.today()
if target_date.day <= 8:
    target_date = target_date - dateutil.relativedelta.relativedelta(months=1)

date_from = f'01.{target_date.month}.{target_date.year}'
date_to = f'{monthrange(target_date.year,target_date.month)[1]}.{target_date.month}.{target_date.year}'

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


def save_dialog(file_name, loops=10, duration=5):
    import win32com.client
    shell = win32com.client.Dispatch("WScript.Shell")
    time.sleep(1)
    shell.SendKeys(r"{ENTER}")
    attempts = 0
    while (attempts < loops):
        print(f"loop {attempts}")
        time.sleep(duration)
        if shell.AppActivate('Save As'):
            shell.SendKeys(r"{TAB}{TAB}{TAB}{TAB}{TAB}")
            shell.SendKeys(file_name)
            shell.SendKeys(r"{ENTER}")
            time.sleep(1)
            if shell.AppActivate('Confirm Save As'):
                shell.SendKeys(r"y")
            attempts = loops

        attempts = attempts+1


def close_excel_by_force():
    import psutil

    for proc in psutil.process_iter():
        if proc.name() == "excel.exe" or proc.name() == "EXCEL.EXE":
            proc.kill()


def fix_excel(file_name):
    import openpyxl
    ss = openpyxl.load_workbook(file_name)
    ss_sheet = ss['Sheet1']
    ss_sheet.title = 'Sheet1'
    ss.save(file_name)
    ss.close()


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

# %%


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
    session.findById("wnd[0]/usr/ctxtKOSTL-LOW").text = "E710000000"
    session.findById("wnd[0]/usr/ctxtKOSTL-HIGH").text = "E889999999"
    session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = date_from
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = date_to
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").caretPosition = 10
    # Get Clase Costo into clipboard
    pyperclip.copy(cf.clase_coste_rpd)
    session.findById("wnd[0]/usr/btn%_KSTAR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/AUTOME"
    session.findById("wnd[0]/usr/btnBUT1").press()
    session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "99999999"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/usr/radRB_OTHERS").setFocus()
    session.findById("wnd[1]/usr/radRB_OTHERS").select()
    session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "10"
    # Workaround to save as SAP dialog windo is not displayed
    save_dialog(file_name=file_name)
    time.sleep(15)
    close_excel_by_force()
    fix_excel(file_name)


def download_kob1_rpd():
    file_name = cf.path_data_local+r'\RPD\kob1_rpd.xlsx'
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKOB1"
    session.findById("wnd[0]").sendVKey(0)
    try:
        session.findById(
            "wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "VT01"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass
    session.findById("wnd[0]/usr/ctxtAUFNR-LOW").text = "1"
    session.findById("wnd[0]/usr/ctxtAUFNR-HIGH").text = "999999999999"
    session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = date_from
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = date_to
    # Get Clase Costo into clipboard
    pyperclip.copy(cf.clase_coste_rpd)
    session.findById("wnd[0]/usr/btn%_KSTAR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtP_DISVAR").setFocus()
    session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 7
    session.findById("wnd[0]/usr/btnBUT1").press()
    session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "99999999"
    session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[43]").press()
    session.findById("wnd[1]/usr/radRB_OTHERS").setFocus()
    session.findById("wnd[1]/usr/radRB_OTHERS").select()
    session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "10"
    # Workaround to save as SAP dialog windo is not displayed
    save_dialog(file_name=file_name)
    time.sleep(15)
    close_excel_by_force()
    fix_excel(file_name)


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
    session.findById(
        "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/radRB_OTHERS").setFocus()
    session.findById("wnd[1]/usr/radRB_OTHERS").select()
    session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "10"
    save_dialog(file_name=file_name, loops=20, duration=10)
    time.sleep(30)
    close_excel_by_force()
    fix_excel(file_name)


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
    session.findById(
        "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById(
        "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/radRB_OTHERS").setFocus()
    session.findById("wnd[1]/usr/radRB_OTHERS").select()
    session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "10"
    save_dialog(file_name=file_name, loops=20, duration=10)
    time.sleep(10)
    logout()
    close_excel_by_force()
    time.sleep(10)
    close_excel_by_force()
    fix_excel(file_name)


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
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/usr/radRB_OTHERS").setFocus()
    session.findById("wnd[1]/usr/radRB_OTHERS").select()
    session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "10"
    # Workaround to save as SAP dialog windo is not displayed
    save_dialog(file_name=file_name, loops=20, duration=10)
    time.sleep(30)
    logout()
    close_excel_by_force()
    time.sleep(10)
    close_excel_by_force()
    fix_excel(file_name)

# %%


def download_ksb1_p15():
    file_name = r'\p15_ksb1_me.xlsx'
    fromcostcenter = '4000000'
    tocostcenter = '6999999'
    costelementgroup = 'H73400'
    layout = '/TSCAMMEV1.0'
    datapath = cf.path_data_local
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
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/usr/radRB_OTHERS").setFocus()
    session.findById("wnd[1]/usr/radRB_OTHERS").select()
    session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
    session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "10"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = datapath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    time.sleep(30)
    close_excel_by_force()


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

# %%


def main():
    try:
        logout()
        login_p11()
        download_ksb1_p11()
        logout()
        login_p15()
        download_ksb1_p15()
        download_koc2_p15()
        logout()
        login_rpd()
        download_kob1_rpd()
        download_ksb1_rpd()
        download_mara_rpd()
        download_t023t_rpd()
        login_rpd()
        download_me2n_rpd()
        logout()
    except Exception as ex:
        print(ex)
        logout()


# %%
if __name__ == '__main__':
    main()
