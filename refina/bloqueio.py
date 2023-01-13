import win32com.client as win32
import subprocess
import pandas as pd

path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
subprocess.Popen(path) #Localiza o 'SAP GUI' e abre o programa
SapGuiAuto = win32.GetObject("SAPGUI")
if not type(SapGuiAuto) == win32.CDispatch :
                print("Login incorreto")

application = SapGuiAuto.GetScriptingEngine
connection = application.OpenConnection("Nome da conexão do SAP", True) #Seleciona a conexão com o banco, deve constar o nome completo do acesso 

session = connection.Children(0)
session.findById("wnd[0]/tbar[0]/okcd").text = "Bloquear"
session.findById("wnd[0]").sendVKey (0)
session.findById("wnd[0]/tbar[1]/btn[17]").press()
session.findById("wnd[1]/usr/txtENAME-LOW").text = "teste"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = (4)
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press()

df = pd.read_excel(r'C:\\Users\\teste\\bloqueio\\bloqueio.xlsx', sheet_name="Planilha1", usecols=[0])
for planilha in df:
    if "Ordens" in planilha:
        df.to_clipboard(excel=True, sep=None, index=False, header=None)

session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/tbar[1]/btn[8]").press()
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell (-1,"")
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll()
session.findById("wnd[0]/tbar[1]/btn[17]").press()
session.findById("wnd[0]/tbar[1]/btn[39]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/tbar[1]/btn[8]").press()
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell (-1,"")
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll()
session.findById("wnd[0]/tbar[1]/btn[48]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/tbar[1]/btn[8]").press()