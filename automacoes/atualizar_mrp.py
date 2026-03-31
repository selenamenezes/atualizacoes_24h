import os
import shutil
from time import sleep
import win32com.client as win32
import traceback
from datetime import datetime
import locale
from dotenv import load_dotenv

load_dotenv()

def atualizar_mrp():
    os.system("taskkill /im excel.exe /f")
    sleep(3)
    pasta = r"C:\BD\MRP_PLANEJAMENTO"
    pasta_centro_atualizacao = os.getenv('PASTA_REP_DASH_MRP')
    base_mrp_plan_v2 = r"C:\BD\MRP_PLANEJAMENTO\BASE_MRP_PLANEJAMENTO_V2.xlsx"
    mrp_plan_v2 = r"C:\BD\MRP_PLANEJAMENTO\MRP_PLAN_V2.xlsm"

    arquivos = os.getenv('LISTA_ARQUIVOS_MRP')

    consultas_base_mrp = [
            "Consulta - ZMM94", "Consulta - ZMMT0003", "Consulta - ZMM208", "Consulta - MB52",
            "Consulta - BASE_CONTRATOS", "Consulta - BASE_PEDIDOS", "Consulta - ESTOQUE_MED_NOVO",
            "Consulta - BASE_MATERIAIS", "Consulta - BASE_MRP", "Consulta - BASE_DEPOSITOS",
            "Consulta - ZMM098", "Consulta - BASE_ZPS047", "Consulta - BASE_ZPS60_V2",
            "Consulta - BASE OBRAS_PO_PLA", "Consulta - BASE_MRP_CONCRETO", "Consulta - ZMM017",
            "Consulta - MB51_ENTREGA_EFET_2026", "Consulta - MB51_ENTRADA_CD_2026",
            "Consulta - BASE_MRP_TELECOM"
        ]

    consultas_mrp_plan = [
            'Consulta - BASE_MRP', 'Consulta - BASE_MRP_CONCRETO', 'Consulta - BASE_PEDIDOS',
            'Consulta - BASE_CONTRATOS', 'Consulta - BASE_MB52', 'Consulta - ANÁLISE DE CONTRATOS',
            'Consulta - BASE_FILTRO_PEDIDOS', 'Consulta - BASE_FILTRO_CONTRATOS',
            'Consulta - ESTOQUE OPERAÇÃO', 'Consulta - OBS', 'Consulta - BASE_ZPS60_V2',
            'Consulta - RESUMO MB52', 'Consulta - CONSUMO_12_MESES'
        ]
    

    def atualizar(planilha, consultas):
        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        wb = None
        try:
            wb = excel.Workbooks.Open(planilha)
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

    separador(lista_arquivos=arquivos, pasta_destino=pasta)
    separador(lista_arquivos=arquivos, pasta_destino=pasta_centro_atualizacao)
    atualizar(base_mrp_plan_v2, consultas_base_mrp)
    atualizar(mrp_plan_v2, consultas_mrp_plan)
    sleep(10)