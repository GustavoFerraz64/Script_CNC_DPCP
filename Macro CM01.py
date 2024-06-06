import pandas as pd
import win32com.client
import time
import sys
import csv
import os
import traceback

"""
Script para extrair dados da transação CM01, ler todos os arquivos da extensão .CAM na pasta \\srvflseng01\Dados\DobraCorte\CEFH-140\Ca Files e
gerar um relatório informando todos os materias que constam na pasta, mas não estão nos dados gerados pela transação CM01

Instruções: 
O usuário deve estar logado em ambiente TPR no SAP, iniciar o script, e um relatório será gerado na pasta \\srvflseng01\Dados\DobraCorte\CEFH-140\CEFH_ROBO_NOVO.xlsx
O arquivo CEFH_ROBO_NOVO.xlsx deve estar fechado, antes de iniciar o script. Caso contrário, irá dar erro.

@autor: Gustavo Nunes Ferraz
@data : 06/06/2024
@departamento: DPCP
@modificado: 06/06/2024
"""

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
        self.caminho_arquivos = r'\\srvflseng01\Dados\DobraCorte\CEFH-140\Ca Files'
        self.caminho_cm01 = r'C:\Temp' #Pasta onde será salvo o txt extraído da transação CM01 do SAP
        self.arquivos_cam = []
        
    #Função destinada a criar uma lista de todos os arquivos .CAM da pasta
    def ler_pasta(self):
        for arquivo in os.listdir(self.pasta):
            if arquivo.endswith('.CAM'):
                self.arquivos_cam.append(arquivo)

    #Função destinada a executar a transação CM01 do SAP, e extrair os dados em CSV.
    def extrai_dados_cm01(self):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/ncm01"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/txt[35,3]").text = "cefh-140"
        self.session.findById("wnd[0]").sendVKey(0)

        #Selecionar todas as  caixas de seleção da transação
        i = 7
        while True:
            try:
                if self.session.findById(f"wnd[0]/usr/chk[1,{i}]") is not None: #Caso retornar FALSE, significa que não há mais caixas a serem selecionadas, e a iteração interrompe.
                    self.session.findById(f"wnd[0]/usr/chk[1,{i}]").selected = True
                    i += 1
            except:
                break

        #Fazer o download do arquivo
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/tbar[1]/btn[20]").press()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.caminho_cm01
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = r'cm01.txt'
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    #Função destinada para ler o arquivo gerado pela CM01, além de tratar os dados
    def ler_dados_cm01(self):
        self.arquivo_cm01= pd.read_csv(r'C:\Users\NUNESFERRAZGUSTAVO\OneDrive - TK Elevator\Área de Trabalho\Script CNC Eduardo\teste.txt', 
                            encoding='latin-1', 
                            sep='\t', 
                            skiprows= 4, 
                            quoting=csv.QUOTE_NONE)
        
        self.arquivo_cm01['Material'] = self.arquivo_cm01['Material'].str.replace('-','')
        self.arquivo_cm01['Material'] = self.arquivo_cm01['Material'].str.replace('.','')
        self.arquivo_cm01 = self.arquivo_cm01[['Dia', 'Material']]
        self.arquivo_cm01.drop(axis=0, index=0)
    
    #Função destinada a ler todos os arquivos com extenção .COM na pasta \\srvflseng01\Dados\DobraCorte\CEFH-140\Ca Files. Realiza também o tratamento de dados.
    def ler_arquivos_pasta(self):

        for arquivo in os.listdir(self.caminho_arquivos):
            if arquivo.endswith('.CAM'):
                self.arquivos_cam.append(arquivo)

        self.arquivos_cam = pd.Series(self.arquivos_cam) #Transforma a lista em series para tratar os dados
        self.arquivos_cam = self.arquivos_cam.str.replace('.CAM','') #Tira o .CAM nos nomes dos arquivos

    def gerar_df_final(self):
        df_resultado = self.arquivo_cm01[~self.arquivo_cm01['Material'].isin(self.arquivos_cam)]
        df_resultado = df_resultado.drop(axis=0, index=0)
        df_resultado.to_excel(r'\\srvflseng01\Dados\DobraCorte\CEFH-140\CEFH_ROBO_NOVO\RELATORIO_ROBO_CEFH.xlsx', index=False)

def main():
    try:
        print("Script iniciado.")
        print("Gerando conexão com o SAP")
        cm01 = CM01()
        print("Extraindo os dados da transação CM01 no SAP.")
        cm01.extrai_dados_cm01()
        print("Lendo os arquivos com extensão .CAM na pasta \\srvflseng01\Dados\DobraCorte\CEFH-140\Ca Files")
        cm01.ler_arquivos_pasta()
        print("Lendo o arquivo gerado pela CM01")
        cm01.ler_dados_cm01()
        print("Compilando os dados e gerando o relatório final.")
        cm01.gerar_df_final()
        os.remove(r'C:\Temp\cm01.txt') #Deletar o arquivo txt extraido pelo SAP. Operação necessária para que o script possa ser rodado novamente no futuro.
        print("Script finalizado. Encerrando o programa.")
        time.sleep(5)
    except Exception as e:
        print(e)
        traceback.print_exc()

if __name__ == '__main__':
    main()