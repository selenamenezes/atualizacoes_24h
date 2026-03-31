import pyautogui as pg
import os
import pandas as pd
import keyboard as kb
from time import sleep
import glob
import pygetwindow as gw
import traceback
import win32com.client as win32
import shutil

PLANILHA_PLA = r"C:\Users\b621314\OneDrive - IBERDROLA S.A\USO INTERNO - NULG\REPOSITORIO\CENTRO_ATUALIZACOES\MRP_PLANEJAMENTO\BASE_PLA.xlsx"
CONSULTAS = ["Consulta - ZCTOBRAS_PLA", "Consulta - BASE_PLA"]
PASTA = r"C:\BD\MRP_PLANEJAMENTO"

def extrair_base_pla():
    sleep(5)
    pg.press('win')
    sleep(3)
    kb.write('edge')
    sleep(3)
    kb.press('enter')
    sleep(5)
    pg.click(x=542, y=50)
    sleep(2)
    kb.write("http://10.211.4.147/DESEMPENHO-NPER/")
    sleep(1)
    kb.press('enter')
    sleep(30)
    pg.click(x=1359, y=98)
    sleep(20)
    pg.click(x=517, y=395)
    sleep(90)
    pg.click(x=847, y=821)
    sleep(5)
    pg.click(x=1336, y=411)
    sleep(3)
    pg.click(x=1323, y=389)
    sleep(3)
    pg.click(x=1259, y=407)
    sleep(3)
    pg.click(x=1135, y=796)

def copiar_notas():
    pasta = r"C:\Users\b621314\Downloads"
    arquivos = glob.glob(os.path.join(pasta, '*'))
    base_pla = max(arquivos, key=os.path.getmtime)

    novo_nome = os.path.join(pasta, "PLA.xlsx")

    if os.path.exists(novo_nome):
        print("Usando arquivo existente PLA.xlsx")
    else:
        nome_arquivo = os.path.basename(base_pla)
        _, extensao = os.path.splitext(base_pla)

        if nome_arquivo.lower().startswith("data") and extensao == ".xlsx":
            os.rename(base_pla, novo_nome)
            print(f"Arquivo renomeado para {novo_nome}")
        else:
            print("Arquivo ignorado: não começa com 'data' ou não é .xlsx")
            return  
        
    df = pd.read_excel(novo_nome)
    col_notas = df['NOTA'].dropna()
    col_notas.to_clipboard(index=False, header=False)

    shutil.copy(novo_nome, PASTA)
    print(f"Arquivo copiado para {PASTA}")


def gerar_projetos(script=r"\\Client\C$\BD\projetos_pla_v2.vbs", ambiente="05.07"):
        max_tentativas = 3
        sucesso = False

        for tentativa in range(1, max_tentativas + 1):
            print(f"Tentativa {tentativa} de executar {os.path.basename(script)}...")

            try:
                    pg.press('win')
                    sleep(3)
                    pg.write("sap")
                    sleep(1)
                    kb.press("enter")
                    sleep(10)  # ideal 300 (5 min)

                    def bring_window_to_front_partial(title_part: str):
                        windows = gw.getAllTitles()
                        matches = [w for w in windows if title_part.lower() in w.lower()]
                        if matches:
                            window = gw.getWindowsWithTitle(matches[0])[0]
                            if window.isMinimized:
                                window.restore()
                            window.activate()
                            window.moveTo(100, 100)
                            return True
                        return False

                    if not bring_window_to_front_partial("SAP Logon Pad"):
                        print("SAP não encontrado, tentando novamente...")
                        continue  

                    sleep(3)
                    pg.click(x=637, y=150)
                    kb.press_and_release("ctrl + a")
                    sleep(1)
                    kb.write(ambiente)
                    sleep(5)
                    kb.press_and_release('enter')
                    sleep(5)
                    kb.press('alt')
                    kb.press('F12')
                    kb.release('F12')
                    kb.release('alt')
                    sleep(3)

                    for i in range(8):
                        kb.press_and_release('down')
                        sleep(0.5)

                    kb.press_and_release('enter')
                    sleep(1)

                    for i in range(6):
                        kb.press_and_release('tab')
                        sleep(1)

                    sleep(5)
                    kb.press('ctrl')
                    kb.press('a')
                    sleep(1)
                    kb.release('ctrl')
                    kb.release('a')
                    sleep(3)
                    kb.press_and_release('delete')
                    kb.write(script)
                    sleep(1)
                    kb.press_and_release('tab')
                    sleep(1)
                    kb.press_and_release('tab')
                    sleep(1)
                    kb.press_and_release('tab')
                    sleep(1)
                    kb.press_and_release('enter')
                    sleep(3)
                    kb.press_and_release('enter')

                    print(f"executando {script}")
                    sucesso = True
                    break 

            except Exception as e:
                    print(f"Erro na tentativa {tentativa} ao executar {script}: {e}")

        if not sucesso:
            print(f"Falhou após {max_tentativas} tentativas. Pulando execução de {script}.")
            return
    
def atualizar_excel_pla():
    os.system("taskkill /im excel.exe /f")
    sleep(5)
    excel = win32.DispatchEx("Excel.Application")
    wb = None
    try:
        wb = excel.Workbooks.Open(PLANILHA_PLA)
        for consulta in CONSULTAS:
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
        print(f"Erro ao abrir/atualizar planilha {PLANILHA_PLA}: {e}")
        print(traceback.format_exc())
    finally:
        if wb:
            wb = None
        excel.Quit()
        excel = None
        
#extrair_base_pla()
copiar_notas()
gerar_projetos()
sleep(100)
atualizar_excel_pla()