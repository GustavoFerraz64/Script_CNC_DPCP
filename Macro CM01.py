import pandas as pd
import win32com.client
import time
import sys
import csv

class CM01:
    
    def __init__(self):
        """Pega a aplicação COM do SAP e a primeira sessão para uso no script
           Testa se o SAP está aberto e se há um usuário logado
           
           """
        try: #Garante que o SAP está aberto, caso contrário encerra o programa
        #Pega a aplicação COM do SAP e a primeira sessão para uso no script
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            self.application = self.SapGuiAuto.GetScriptingEngine
            self.connection = self.application.Children(0)
            self.sessions = self.connection.Children
            self.session = self.sessions[0]
        except:
            print("SAP não está aberto, o programa será finalizado")
            time.sleep(2)
            sys.exit(0)
        #Testa se há um usuário logado no SAP, caso não haja, encerra o programa
        self.usuario = self.session.Info.User
        if self.usuario == '':
            print("SAP não está logado, o programa será finalizado")
            time.sleep(2)
            sys.exit(0)
        
        #Define o caminho onde serão armazenados os arquivos
        self.path = r'\\srvflseng01\dados\DobraCorte'
        
        def cm01(self):