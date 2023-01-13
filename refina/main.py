import pyautogui
import win32com.client as win32
import subprocess
import time
import os
import sys
import webbrowser
import pandas as pd

while True : #Condição 'TRUE' para manter o SAP aberto, enquanto o script é executado 

            opcao = pyautogui.confirm("Script para extrair informações , selecione um das variantes e aguarde a execução:", buttons=['1', '2', 'Navegador', 'Sair']) #Tela para seleção de opções   

            pyautogui.PAUSE = 0.5

            if opcao == "Sair":
                sys.exit(opcao)  #Encerra o script

            if opcao == "Navegador":

                webbrowser.open("https://www.google.com.br")
                break #Abre o navegador padrão na página correspondente

            path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
            subprocess.Popen(path) #Localiza o 'SAP GUI' e abre o programa
            time.sleep(5)

            SapGuiAuto = win32.GetObject("SAPGUI")

            if not type(SapGuiAuto) == win32.CDispatch :
                print("Login incorreto")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.OpenConnection("Nome da conexão do SAP", True) #Seleciona a conexão com o banco, deve constar o nome completo do acesso
            time.sleep(3)
            session = connection.Children(0) #Cria a sessão de conexão com o SAP
            time.sleep(2)

            if opcao == "1" :
                print("=============Iniciando script SAP '1'================")
                session.findById("wnd[0]/tbar[0]/okcd").text = "1"
                session.findById("wnd[0]").sendVKey (0)
                session.findById("wnd[0]/tbar[1]/btn[17]").press()
                session.findById("wnd[1]/usr/txtV-LOW").text = "1_variante"
                session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
                session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
                session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = (0)
                session.findById("wnd[1]/tbar[0]/btn[8]").press()
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem ("&XXL")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\teste"
                session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
                session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = (97)
                session.findById("wnd[1]/tbar[0]/btn[0]").press()                

                caminhoArquivos = r"C:\\Users\\teste"
                listaArquivos = os.listdir(caminhoArquivos)
                listaCaminhoEArquivoExcel = [caminhoArquivos + '\\' + arquivo for arquivo in listaArquivos if arquivo[-4:] == 'XLSX']
                #Verifica no diretorio se consta o arquivo 'XLSX', fazendo a verificação do nome da extensão
                dadosArquivo = pd.DataFrame() #Cria o DataFrame dos arquivos para a cópia 

                for arquivo in listaCaminhoEArquivoExcel: #Seleciona as linhas e colunas para cópia dos dados 
                    dados = pd.read_excel(arquivo)
                    dadosArquivo = dadosArquivo.append(dados)
                dadosArquivo.to_excel(r"C:\\Users\\teste\\planilha_testes_1.xlsx", index=False) #Converte o DataFrame para arquivo de Excel

                excel = win32.DispatchEx('Excel.Application') #Acessa a aplicação do Excel e altera o nome da Sheet
                wb = excel.Workbooks.Open("C:\\Users\\teste")
                wb.Worksheets(1).Name = '1'
                excel = win32.Dispatch('Excel.Application') #Fecha a aplicação Excel  

                time.sleep(10)
                excel.Quit()
                time.sleep(10)

                if os.path.exists("C:\\Users\\teste\\EXPORT.XLSX"):
                    os.remove("C:\\Users\\teste\\EXPORT.XLSX")
                else:
                    print('Arquivo já deletado') #Condição para excluir o arquivo 'EXPORT', caso já exista informa que o arquivo já foi deletado

                excel = win32.Dispatch('Excel.Application') #Fecha a aplicação Excel 

                time.sleep(10)
                excel.Quit()               

                print("============Fim do script '1'=================")

            if opcao == "2" :
                print("=============Iniciando script SAP '2'================")

                session.findById("wnd[0]/tbar[0]/okcd").text = "2"
                session.findById("wnd[0]").sendVKey (0)
                session.findById("wnd[0]/tbar[1]/btn[17]").press()
                session.findById("wnd[1]/usr/txtV-LOW").text = "2_variante"
                session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
                session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
                session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = (0)
                session.findById("wnd[1]/tbar[0]/btn[8]").press()
                session.findById("wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH").press()

                df = pd.read_excel(r'C:\\Users\\teste\\planilha_testes_1.xlsx, sheet_name="1", usecols=[0])
                for planilha in df:
                    if "Nota" in planilha:
                        df.to_clipboard(excel=True, sep=None, index=False, header=None) #Le o arquivo gerado na sessão, copía a coluna de 'Nota' e cola no SAP para realizar a consulta  

                session.findById("wnd[1]/tbar[0]/btn[24]").press()
                session.findById("wnd[1]/tbar[0]/btn[8]").press()
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = ("HEYTO")
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem ("&XXL")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\teste\\2"
                session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
                session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = (83)
                session.findById("wnd[1]/tbar[0]/btn[0]").press()

                caminhoArquivos = r"C:\\Users\\teste\\2"
                listaArquivos = os.listdir(caminhoArquivos)
                listaCaminhoEArquivoExcel = [caminhoArquivos + '\\' + arquivo for arquivo in listaArquivos if arquivo[-4:] == 'XLSX']

                dadosArquivo = pd.DataFrame()
                for arquivo in listaCaminhoEArquivoExcel:
                    dados = pd.read_excel(arquivo)
                    dadosArquivo = dadosArquivo.append(dados)

                dadosArquivo.to_excel(r"C:\\Users\\teste\\2\\planilha_testes_2.xlsx", index=False)

                excel = win32.DispatchEx('Excel.Application')
                wb = excel.Workbooks.Open("C:\\Users\\teste\\2\\planilha_teste_2.xlsx")
                wb.Worksheets(1).Name = '2'

                excel = win32.Dispatch('Excel.Application')
                time.sleep(10)
                excel.Quit()
                time.sleep(10)

                if os.path.exists("C:\\Users\\teste\\2\\EXPORT.XLSX"):
                    os.remove("C:\\Users\\teste\\2\\EXPORT.XLSX")
                else:
                    print('Arquivo já deletado')

                excel = win32.Dispatch('Excel.Application')
                time.sleep(10)
                excel.Quit()