import os
import shutil
from time import sleep
import win32com.client as win32
import traceback
from datetime import datetime
import config
def atualizar_cobertura():

    os.system("taskkill /im excel.exe /f")

    pasta = r"C:\BD\DASH_COBERTURA_PLANEJADO"
    pasta_repositorio = config.PASTA_REP_DASH_COBERT
    bd_cobertura = r"C:\BD\DASH_COBERTURA_PLANEJADO\BD_DASH_COBERTURA_PLANEJ_V5_beta.xlsx"
    dash_cobertura = r"C:\BD\DASH_COBERTURA_PLANEJADO\DASH DE COBERTURA DE MAT PLAN_V4.xlsx"

    arquivos = config.ARQUIVOS_COBERTURA_BD

    consultas_bd = ["Consulta - CARTEIRA_OBRAS", "Consulta - CARTEIRA ORDENS", "Consulta - MB52", "Consulta - ZMM94",
                    "Consulta - BASE_ZPS047", "Consulta - BASE_CONTRATOS", "Consulta - BASE_PEDIDOS", "Consulta - BASE_ZPS60_V2",
                    "Consulta - BASE_ORÇAMENTO", "Consulta - PORTE DE OBRA", "Consulta - T_DEMANDA_EPS", 
                    "Consulta - T_DEMANDA_ORDENS_V2",
                    "Consulta - T_COBERTURA_PROJETOS"
                    ]

    consultas_dash = ["Consulta - T_COBERTURA_PROJETOS", "Consulta - T_DEMANDA_EPS", "Consulta - CARTEIRA DE OBRAS", "Consulta - T_DEMANDA_EPS_GERAL",
                    "Consulta - T_DEMANDA_EPS_CONGELADA", "Consulta - ENTREGAS DE CONCRETOS", "Consulta - BASE_ZPS047", "Consulta - T_ORÇAMENTO_ORDENS",
                    "Consulta - BASE_MB52", "Consulta - Ações-Obs", "Consulta - BASE_ZMM94", "Consulta - BASE AÇÕES"]

    def separador(lista_arquivos, pasta_destino):
        os.makedirs(pasta_destino, exist_ok=True)

        for caminho_arquivo in lista_arquivos:
            try:
                nome_arquivo = os.path.basename(caminho_arquivo)
                destino_final = os.path.join(pasta_destino, nome_arquivo)
                shutil.copy2(caminho_arquivo, destino_final)
                print(f"Copiado: {caminho_arquivo} → {destino_final}")
            except Exception as e:
                print(f"Erro ao copiar {caminho_arquivo}: {e}")

    def atualizar(planilha, consultas):
        excel = win32.DispatchEx("Excel.Application")
        
        wb = None
        try:
            wb = excel.Workbooks.Open(planilha)
            print(f"Atualizando: {os.path.basename(planilha)}")
            for consulta in consultas:
                print(f"Atualizando Consulta: {consulta}")
                try:
                    connection = next((c for c in wb.Connections if c.Name == consulta), None)
                    if connection and connection.Type == 1:
                        oledb = connection.OLEDBConnection
                        oledb.BackgroundQuery = False
                        oledb.Refresh()
                        print(f"\tConsulta '{consulta}' atualizada!")
                    else:
                        print(f"Conexão '{consulta}' não encontrada ou inválida.")
                except Exception as e:
                    print(f"Erro ao atualizar '{consulta}': {e}")
                    print(traceback.format_exc())
            wb.Close(SaveChanges=1)
        except Exception as e:
            print(f"Erro ao abrir/atualizar planilha {planilha}: {e}")
            print(traceback.format_exc())
        finally:
            if wb:
                wb = None
            excel.Quit()
            excel = None

    #separador(lista_arquivos=arquivos, pasta_destino=pasta)
    #separador(lista_arquivos=arquivos, pasta_destino=pasta_repositorio)
    #sleep(5)
    atualizar(bd_cobertura, consultas_bd)
    atualizar(dash_cobertura, consultas_dash)
    sleep(10)
    shutil.copy2(bd_cobertura, config.PASTA_REP_DASH_COBERT)
    shutil.copy2(dash_cobertura, config.PASTA_REP_DASH_COBERT)

    DESTINO_COBERTURA = config.DESTINO_COBERTURA

    data_atual = datetime.now()
    ano = data_atual.strftime("%Y")
    mes_num = data_atual.strftime("%m")
    mes_nome = data_atual.strftime("%B").upper()
    dia = data_atual.strftime("%Y%m%d")
    pasta_mes = f"{mes_num}. {mes_nome}"

    destino_final = os.path.join(DESTINO_COBERTURA, ano, pasta_mes, dia)
    os.makedirs(destino_final, exist_ok=True)
    shutil.copy2(dash_cobertura, destino_final)

atualizar_cobertura()
