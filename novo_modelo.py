from datetime import datetime
import keyboard
import time
import pyautogui
import pygetwindow as gw
import ctypes
import threading
import schedule
import os
import shutil
from automacoes.atualizar_mrp import atualizar_mrp
from automacoes.atualizar_cobertura import atualizar_cobertura
import locale
import pathlib
from dotenv import load_dotenv

load_dotenv()

FULL_BASES_TESTE_PT1 = os.getenv('FULL_BASES_PT1')
FULL_BASES_TESTE_PT2 = os.getenv('FULL_BASES_PT2')
BASES_PERIODICAS = os.getenv('BASES_PERIODICAS')
MED_CCS = os.getenv('MED_CCS')
MB52_r3 = os.getenv('MB52_r3')
BASE_NF = os.getenv('BASE_NF')
BASE_MRP_PLAN_V2 = os.getenv('BASE_MRP_PLAN_V2')
MRP_PLAN = os.getenv('MRP_PLAN')
AMBIENTE_R3 = os.getenv('AMBIENTE_r3')
AMBIENTE_CCS = os.getenv('AMBIENTE_CCS')

def manter_ativo():
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002 | 0x00000001)

def extrair_sap(script, ambiente):
    max_tentativas = 3
    sucesso = False

    for tentativa in range(1, max_tentativas + 1):
        print(f"Tentativa {tentativa} de executar {os.path.basename(script)}...")

        try:
            pyautogui.press('win')
            time.sleep(3)
            pyautogui.write("sap")
            time.sleep(1)
            keyboard.press("enter")
            time.sleep(300)  # ideal 300 (5 min)

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

            time.sleep(3)
            pyautogui.click(x=637, y=150)
            keyboard.press_and_release("ctrl + a")
            time.sleep(1)
            keyboard.write(ambiente)
            time.sleep(5)
            keyboard.press_and_release('enter')
            time.sleep(5)
            keyboard.press('alt')
            keyboard.press('F12')
            keyboard.release('F12')
            keyboard.release('alt')
            time.sleep(3)

            for i in range(8):
                keyboard.press_and_release('down')
                time.sleep(0.5)

            keyboard.press_and_release('enter')
            time.sleep(1)

            for i in range(6):
                keyboard.press_and_release('tab')
                time.sleep(1)

            time.sleep(5)
            keyboard.press('ctrl')
            keyboard.press('a')
            time.sleep(1)
            keyboard.release('ctrl')
            keyboard.release('a')
            time.sleep(3)
            keyboard.press_and_release('delete')
            keyboard.write(script)
            time.sleep(1)
            keyboard.press_and_release('tab')
            time.sleep(1)
            keyboard.press_and_release('tab')
            time.sleep(1)
            keyboard.press_and_release('tab')
            time.sleep(1)
            keyboard.press_and_release('enter')
            time.sleep(3)
            keyboard.press_and_release('enter')

            print(f"{datetime.now()} - executando {script}")
            sucesso = True
            break 

        except Exception as e:
            print(f"Erro na tentativa {tentativa} ao executar {script}: {e}")

    if not sucesso:
        print(f"Falhou após {max_tentativas} tentativas. Pulando execução de {script}.")
        return

def verificar_janela(janela_aberta):
    try:
        janela = pyautogui.getWindowsWithTitle(janela_aberta)
        for _ in range(30):
            if janela:
                print(f"{janela_aberta} esta aberto")
                break
        else:
            print(f"{janela_aberta} nao esta aberto")
    except IndexError:
        print("erro ao abrir buscar janela")

sisbase = os.getenv('BASES_SISBASE')
cobertura = os.getenv('BASES_COBERTURA')
concreto = os.getenv('BASES_CONCRETO')
planejamento = os.getenv('BASES_PLANEJAMENTO')
estoque = os.getenv('BASES_ESTOQUE')
transito = os.getenv('BASES_TRANSITO')
sucata = os.getenv('BASES_SUCATA')
medidores = os.getenv('BASES_MEDIDORES')
nf = os.getenv('BASES_NF')
geral_monitorada = os.getenv('GERAL_MONITORADA')
pasta_sisbase = os.getenv('PASTA_SISBASE')
pasta_concreto = os.getenv('PASTA_CONCRETO')
pasta_cobertura = os.getenv('PASTA_COBERTURA')
pasta_mrp_plan = os.getenv('PASTA_MRP_PLAN')
pasta_sucata = os.getenv('PASTA_SUCATA')
pasta_transito = os.getenv('PASTA_TRANSITO')
pasta_estoque = os.getenv('PASTA_ESTOQUE')
pasta_medidor_bi = os.getenv('PASTA_MEDIDOR_BI')
pasta_mrp_plan_nuvem = os.getenv('PASTA_MRP_PLAN_NUVEM')
pasta_concreto_bi = os.getenv('PASTA_CONCRETO_BI')
pasta_bd_mrp = os.getenv('PASTA_BD_MRP')
me2m = os.getenv('ME2M')
entregas_cif = os.getenv('ENTREGAS_CIF')
arquivos_especificos = os.getenv('ARQUIVOS_ESPECIFICOS')
historico_zmm94 = os.getenv('HISTORICO_ZMM94')
historico_mb52 = os.getenv('HISTORICO_MB52')
historico_rel_em = os.getenv('HISTORICO_REL_EM')
base_historico_mb52 = os.getenv('BASE_HISTORICO_MB52')
base_historico_zmm94 = os.getnev('BASE_HISTORICO_ZMM94')
base_historico_rel_em = os.getenv('BASE_HISTORICO_REL_EM')

def separador_historico(base, destino):
    data_modif = os.path.getmtime(base)
    data_modificacao = datetime.fromtimestamp(data_modif).strftime('%d.%m.%Y')
    nome_base = os.path.basename(base)
    _, ext = os.path.splitext(nome_base)

    novo_nome = f'{data_modificacao}{ext}'
    destino = os.path.join(destino, novo_nome)

    shutil.copy2(base, destino)

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

def copiar_planilha(origem, destino_pasta):
    if not os.path.exists(origem):
        print("Arquivo não encontrado:", origem)
        return
    
    if not os.path.exists(destino_pasta):
        print("Pasta de destino não encontrada:", destino_pasta)
        return
    
    try:
        nome_arquivo = os.path.basename(origem)
        destino = os.path.join(destino_pasta, nome_arquivo)
        shutil.copy2(origem, destino)
        print(f"Arquivo copiado com sucesso para {destino}")
    except Exception as e:
        print("Erro ao copiar o arquivo:", e)

def movimentar_medidores():
    separador_historico(base=base_historico_rel_em, destino=historico_rel_em)
    separador(lista_arquivos=medidores, pasta_destino=pasta_medidor_bi)

def movimentar_full():
    separador(lista_arquivos=me2m, pasta_destino=pasta_mrp_plan_nuvem)
    separador(lista_arquivos=sisbase, pasta_destino=pasta_sisbase)
    separador(lista_arquivos=concreto, pasta_destino=pasta_concreto)
    separador(lista_arquivos=cobertura, pasta_destino=pasta_cobertura)
    separador(lista_arquivos=planejamento, pasta_destino=pasta_mrp_plan)

def movimentar_periodicas():
    separador(lista_arquivos=sucata, pasta_destino=pasta_sucata)
    separador(lista_arquivos=transito, pasta_destino=pasta_transito)
    separador(lista_arquivos=estoque, pasta_destino=pasta_estoque)
    separador_historico(base=base_historico_mb52, destino=historico_mb52)
    separador_historico(base=base_historico_zmm94, destino=historico_zmm94)
    

def movimentar_nf():
    separador(lista_arquivos=nf, pasta_destino=pasta_concreto_bi)
    separador(lista_arquivos=entregas_cif, pasta_destino=pasta_concreto_bi)
    separador(lista_arquivos=entregas_cif, pasta_destino= os.getenv('PASTA_NUVEM_DASH_MOVIMENTACAO'))

def movimentar_bases_especificas(arquivos, pasta_mrp):
    separador(lista_arquivos=arquivos, pasta_destino=pasta_mrp)

def backup_mrp():
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil')

        arquivos_salvar = ['BASE_MRP_PLANEJAMENTO_V2.xlsx', 'MRP_PLAN_V2.xlsm']
        origem = rf'C:\BD\MRP_PLANEJAMENTO'
        pasta_mrp = os.getenv('PASTA_MRP')

        ano_atual = datetime.now().strftime('%Y')
        mes = f"{datetime.now().strftime('%m')}. {datetime.now().strftime('%b').upper()}"
        dia_execucao = datetime.now().strftime('%Y%m%d')

        pasta_ano = os.path.join(pasta_mrp, ano_atual)
        pasta_mes = os.path.join(pasta_ano, mes)
        pasta_dia = os.path.join(pasta_mes, dia_execucao)

        if not os.path.exists(pasta_ano):
            os.makedirs(pasta_ano)
            print(f'Criada pasta do ano: {pasta_ano}')

        if not os.path.exists(pasta_mes):
            os.makedirs(pasta_mes)
            print(f'Criada pasta do mês: {pasta_mes}')

        if not os.path.exists(pasta_dia):
            os.makedirs(pasta_dia)
            print(f'Criada pasta do dia: {pasta_dia}')

        os.startfile(pasta_dia)

        for arquivo in arquivos_salvar:
            c_origem = os.path.join(origem, arquivo)
            c_destino = os.path.join(pasta_dia, arquivo)

            if os.path.isfile(c_origem):
                shutil.copy2(c_origem, c_destino)
                print(f"Copiado: {arquivo} => {c_destino}")
            else:
                print(f"Arquivo não encontrado: {c_origem}")

def sub_orig():
    c_origem = rf'C:\BD\MRP_PLANEJAMENTO'
    c_destino = os.getenv('PASTA_DASH_MRP_NUVEM')

    arquivos_substituir = ['BASE_MRP_PLANEJAMENTO_V2.xlsx', 'MRP_PLAN_V2.xlsm']

    for arquivo in arquivos_substituir:
        origem_ = os.path.join(c_origem, arquivo)
        destino = os.path.join(c_destino, arquivo)
        if os.path.exists(origem_):
            shutil.copy2(origem_, destino)
            print(rf"MRPs atualizados adicionados no caminho \BASES POWER BI\DASH MRP PLANEJAMENTO")
        else:
            print(f"Arquivo {origem_} em {c_origem} não encontrado")

    destino = os.getenv('PASTA_BD_MRP_NUVEM')

    for arquivo in arquivos_substituir:
        origem = os.path.join(c_origem, arquivo)
        destino_ = os.path.join(destino, arquivo)
        if os.path.exists(origem):
            shutil.copy2(origem, destino_)
            print(rf"MRPs atualizados adicionados no caminho \BD\MRP_PLANEJAMENTO")
        else:
            print(f"Arquivo {origem} em {c_origem} não encontrado")

def verificar_bases():
    try:
        mb52 = pathlib.Path(os.getenv('TXT_MB52'))
        me2m = pathlib.Path(os.getenv('TXT_ME2M_PC'))

        mb52_data = datetime.fromtimestamp(mb52.stat().st_mtime).date()
        me2m_data = datetime.fromtimestamp(me2m.stat().st_mtime).date()

        hoje = datetime.today().date()

        if me2m_data == hoje and mb52_data == hoje:
            print("Divergência de bases")
            return
        else:
            extrair_sap(script=r"C:\BD\MRP_v3.vbs", ambiente='06.01')
            return

    except Exception as e:
        print(f"erro: {e}")

sistema_ativo = threading.Thread(target=manter_ativo, daemon=False)
sistema_ativo.start()

pyautogui.FAILSAFE = False


schedule.every().day.at("07:30").do(lambda:extrair_sap(MB52_r3, AMBIENTE_R3))
schedule.every().day.at("07:50").do(lambda:extrair_sap(MED_CCS, AMBIENTE_CCS))
schedule.every().day.at("08:00").do(movimentar_medidores)

schedule.every().day.at("08:25").do(lambda:extrair_sap(BASES_PERIODICAS, AMBIENTE_R3))
schedule.every().day.at("08:40").do(movimentar_periodicas)

schedule.every().day.at("11:00").do(lambda:extrair_sap(BASES_PERIODICAS, AMBIENTE_R3))
schedule.every().day.at("11:30").do(movimentar_periodicas)

schedule.every().day.at("13:00").do(lambda:extrair_sap(BASES_PERIODICAS, AMBIENTE_R3))
schedule.every().day.at("13:40").do(movimentar_periodicas)

schedule.every().day.at("13:53").do(lambda:extrair_sap(BASE_NF, AMBIENTE_R3))
schedule.every().day.at("14:20").do(lambda:movimentar_nf)

schedule.every().day.at("15:00").do(lambda:extrair_sap(BASES_PERIODICAS, AMBIENTE_R3))
schedule.every().day.at("15:40").do(movimentar_periodicas)

schedule.every().day.at("17:00").do(lambda:extrair_sap(BASES_PERIODICAS, AMBIENTE_R3))
schedule.every().day.at("17:40").do(movimentar_periodicas)

schedule.every().day.at("18:00").do(lambda:extrair_sap(FULL_BASES_TESTE_PT1, AMBIENTE_R3))
schedule.every().day.at("19:10").do(lambda:extrair_sap(FULL_BASES_TESTE_PT2, AMBIENTE_R3))
schedule.every().day.at("19:15").do(verificar_bases)
schedule.every().day.at("21:00").do(lambda:movimentar_full)

schedule.every().day.at("21:10").do(lambda:extrair_sap(BASES_PERIODICAS, AMBIENTE_R3))
schedule.every().day.at("21:40").do(movimentar_periodicas)

try:
    schedule.every().day.at("21:30").do(atualizar_cobertura)
    schedule.every().day.at("04:00").do(lambda:movimentar_bases_especificas(arquivos=arquivos_especificos, pasta_mrp=pasta_bd_mrp))
    schedule.every().day.at("04:20").do(atualizar_mrp) 
    schedule.every().day.at("06:00").do(backup_mrp)
    schedule.every().day.at("06:10").do(sub_orig)
except Exception as e:
    print(f"erro: {e}")

while True:
    schedule.run_pending()
    time.sleep(1);                                                                  