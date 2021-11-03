import win32com.client
import os
import logging
from typing import Optional
import time

class SAPConnection:
    def __init__(self,
    sapPath:str=r'C:/Program Files (x86)/SAP/FrontEnd/SapGui/saplogon.exe',
    user: Optional[str] = None,
    pwd:Optional[str] = None,
    system:Optional[str] = None,
    language:str="EN"
    ):
        logging.info(f"Connected to SAP with user {user} to system {system}")
        not_logged = True
        os.startfile(sapPath)
        while not_logged:
            try:
                SapGuiAuto = win32com.client.GetObject("SAPGUI")
                SAPApp = SapGuiAuto.GetScriptingEngine
                SAPCon = SAPApp.OpenConnection(system, True)
                session = SAPCon.Children(0)
                not_logged = False
            except:
                time.sleep(1)
        session.findById("wnd[0]").iconify()
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = pwd
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = language
        session.findById("wnd[0]").sendVKey (0)
        if session.ActiveWindow.Name == "wnd[1]":
            if 'Multiple Logon' in session.findById("wnd[1]").text:
                session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
                session.findById("wnd[1]").sendVKey (0)
        self.session = session

    def break_connection(self):
        try:
            self.session = SAPCon.Children(0)
            self.session.findById("wnd[0]").Close()
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except:
                pass
        except:
            pass
        try:
            os.system("TASKKILL /F /IM saplogon.exe")
        except:
            pass
