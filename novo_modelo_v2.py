from datetime import datetime
import time
import os
import shutil
import threading
import locale
import pathlib
import ctypes

import pyautogui
import keyboard
import pygetwindow as gw
import schedule

from automacoes.atualizar_mrp import atualizar_mrp
from automacoes.atualizar_cobertura import atualizar_cobertura
import automacoes.config as config


# config
pyautogui.FAILSAFE = False

AMBIENTE_R3 = config.AMBIENTE_R3
AMBIENTE_CCS = config.AMBIENTE_CCS


# utils
def manter_pc_ativo():
    """Evita que o PC entre em modo de suspensão."""
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002 | 0x00000001)


def esperar(segundos, msg=""):
    if msg:
        print(msg)
    time.sleep(segundos)


def copiar_arquivo(origem, destino):
    try:
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        shutil.copy2(origem, destino)
        print(f"copiado: {origem} → {destino}")
    except Exception as e:
        print(f"erro ao copiar {origem}: {e}")


# automacao sap
def abrir_sap():
    pyautogui.press('win')
    esperar(2)
    pyautogui.write("sap")
    esperar(1)
    keyboard.press_and_release("enter")
    esperar(300, "abrindo SAP...")


def focar_janela(titulo_parcial):
    for title in gw.getAllTitles():
        if titulo_parcial.lower() in title.lower():
            win = gw.getWindowsWithTitle(title)[0]
            if win.isMinimized:
                win.restore()
            win.activate()
            win.moveTo(100, 100)
            return True
    return False


def login_sap(ambiente):
    pyautogui.click(x=637, y=150)
    keyboard.send("ctrl+a")
    keyboard.write(ambiente) # inserir ambiente do sap
    esperar(5)
    keyboard.send("enter")

    # abrir opções
    keyboard.send("alt+F12")
    esperar(2)

    for _ in range(8):
        keyboard.send("down")

    keyboard.send("enter")

    for _ in range(6):
        keyboard.send("tab")

    keyboard.send("ctrl+a")
    keyboard.send("delete")


def executar_script(script):
    keyboard.write(script)
    keyboard.send("tab tab tab enter")
    esperar(2)
    keyboard.send("enter")


def extrair_sap(script, ambiente, tentativas=3):
    for tentativa in range(1, tentativas + 1):
        print(f"Tentativa {tentativa}: {script}")

        try:
            abrir_sap()

            if not focar_janela("SAP Logon Pad"):
                print("SAP não encontrado")
                continue

            login_sap(ambiente)
            executar_script(script)

            print(f"Executando: {script}")
            return True

        except Exception as e:
            print(f"erro: {e}")

    print(f"falha após {tentativas} tentativas")
    return False


# arquivos
def copiar_lista(arquivos, destino): # funcao para copiar uma lista de arquivos e colar em uma pasta especifica
    os.makedirs(destino, exist_ok=True)
    for arq in arquivos:
        destino_final = os.path.join(destino, os.path.basename(arq))
        copiar_arquivo(arq, destino_final)


def salvar_historico(arquivo, destino): # funcao para salvar o historico de um arquivo em uma pasta especifica (Ex do nome do arquivo. 12.01.2026.txt)
    data = datetime.fromtimestamp(os.path.getmtime(arquivo)).strftime('%d.%m.%Y')
    nome = f"{data}{os.path.splitext(arquivo)[1]}"
    copiar_arquivo(arquivo, os.path.join(destino, nome))


# movimentacoes
def movimentar_medidores(): # funcoes que unem as funcoes de salvar historico e copiar lista para organizar os arquivos de medidores
    salvar_historico(config.BASE_HISTORICO_REL_EM, config.HISTORICO_REL_EM)
    copiar_lista(config.BASES_MEDIDORES, config.PASTA_MEDIDOR_BI)


def movimentar_full(): # funcao que une as funcoes de salvar historico e copiar lista para organizar os arquivos do script full
    copiar_lista(config.BASES_SISBASE, config.PASTA_SISBASE)
    copiar_lista(config.BASES_CONCRETO, config.PASTA_CONCRETO)
    copiar_lista(config.BASES_COBERTURA, config.PASTA_COBERTURA)
    copiar_lista(config.BASES_PLANEJAMENTO, config.PASTA_MRP_PLAN)


def movimentar_periodicas(): # funcao que une as funcoes de salvar historico e copiar lista para organizar os arquivos de atualizacao periodica
    copiar_lista(config.BASES_SUCATA, config.PASTA_SUCATA)
    copiar_lista(config.BASES_TRANSITO, config.PASTA_TRANSITO)
    copiar_lista(config.BASES_ESTOQUE, config.PASTA_ESTOQUE)

    salvar_historico(config.BASE_HISTORICO_MB52, config.HISTORICO_MB52)
    salvar_historico(config.BASE_HISTORICO_ZMM94, config.HISTORICO_ZMM94)


# backup
def backup_mrp():
    locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil')

    origem = r'C:\BD\MRP_PLANEJAMENTO'
    arquivos = ['BASE_MRP_PLANEJAMENTO_V2.xlsx', 'MRP_PLAN_V2.xlsm']

    hoje = datetime.now()

    pasta = os.path.join(
        config.PASTA_MRP,
        hoje.strftime('%Y'),
        f"{hoje.strftime('%m')}. {hoje.strftime('%b').upper()}",
        hoje.strftime('%Y%m%d')
    )

    os.makedirs(pasta, exist_ok=True)
    os.startfile(pasta)

    for arq in arquivos:
        copiar_arquivo(os.path.join(origem, arq), os.path.join(pasta, arq))


def atualizar_bases_nuvem():
    origem = r'C:\BD\MRP_PLANEJAMENTO'
    arquivos = ['BASE_MRP_PLANEJAMENTO_V2.xlsx', 'MRP_PLAN_V2.xlsm']

    destinos = [
        config.PASTA_DASH_MRP_NUVEM,
        config.PASTA_BD_MRP_NUVEM
    ]

    for destino in destinos:
        for arq in arquivos:
            copiar_arquivo(os.path.join(origem, arq), os.path.join(destino, arq))


# validacoes
def verificar_bases():
    try:
        mb52 = pathlib.Path(config.TXT_MB52)
        me2m = pathlib.Path(config.TXT_ME2M_PC)

        hoje = datetime.today().date()

        if (datetime.fromtimestamp(mb52.stat().st_mtime).date() != hoje or
                datetime.fromtimestamp(me2m.stat().st_mtime).date() != hoje):

            print("bases desatualizadas >> executando SAP")
            extrair_sap(r"C:\BD\MRP_v3.vbs", AMBIENTE_R3)

    except Exception as e:
        print(f"Erro verificação: {e}")


# agendamentos
def configurar_agendamentos():
    schedule.every().day.at("20:00").do(
        lambda: extrair_sap(config.SCRIPT_60_47, AMBIENTE_R3)
    )

    schedule.every().day.at("14:32").do(atualizar_cobertura)
    schedule.every().day.at("04:00").do(atualizar_mrp)
    schedule.every().day.at("06:40").do(backup_mrp)
    schedule.every().day.at("06:45").do(atualizar_bases_nuvem)


# main
if __name__ == "__main__":
    threading.Thread(target=manter_pc_ativo, daemon=True).start()

    configurar_agendamentos()

    print(">> SCRIPT INICIADO <<")

    while True:
        schedule.run_pending()
        time.sleep(1)