# Automação diária (SAP → Bases + Movimentações) — Python

Script de automação para executar rotinas em **SAP via automação de interface** (PyAutoGUI/keyboard) e organizar/backup das bases para consumo analítico (Power BI/dashboards). O fluxo é agendado por horários usando `schedule`.

> Este script foi projetado para ambiente **Windows**, com SAP Logon Pad e atalhos/menus específicos. A dependência do layout/telas do SAP é alta.

## Objetivo
- Abrir o **SAP Logon**.
- Selecionar um **ambiente** (ex.: `06.01`, `R3`, `CCS`).
- Gerar/extrair bases executando rotinas a partir de arquivos/scripts.
- Mover/arquivar arquivos para pastas organizadas.
- Atualizar rotinas auxiliares (`atualizar_mrp`, `atualizar_cobertura`).
- Fazer backup de arquivos do planejamento do MRP em base diária.

## Tecnologias e bibliotecas
- `pyautogui` (interação com teclado/mouse)
- `keyboard` (envio de teclas)
- `pygetwindow` (localizar/jogar janela no foco)
- `schedule` (agendamento por horário)
- `ctypes` (manter processo ativo via SetThreadExecutionState)
- `pandas` não é usado diretamente (este script foca automação/arquivos)
- `locale`, `pathlib`, `os`, `shutil`, `threading`, `time`, `datetime`
- Módulos locais:
  - `automacoes.atualizar_mrp`
  - `automacoes.atualizar_cobertura`
  - `automacoes.config` (configurações e listas de arquivos/pastas)

## Configuração (automacoes/config.py)
O script importa muitas constantes de `automacoes.config as config`.

Em geral, espera que `config.py` forneça:
- Caminhos para bases/arquivos a extrair (`BASES_*`) e seus respectivos scripts/nomes
- Pastas destino (`PASTA_*`)
- Parâmetros para rotinas auxiliares:
  - `atualizar_mrp`
  - `atualizar_cobertura`
- Caminhos/paths para arquivos de monitoramento (`TXT_MB52`, `TXT_ME2M`, etc.)

## Principais funções

### `manter_ativo()`
Evita o sistema “dormir” durante a execução chamando `ctypes.windll.kernel32.SetThreadExecutionState(...)`.

### `extrair_sap(script, ambiente)`
Executa uma sequência de cliques/teclas:
1. Abre o SAP (busca pelo programa via teclado `win` → digita `sap`).
2. Localiza a janela com título contendo **"SAP Logon Pad"**.
3. Seleciona a “opção” do ambiente (`keyboard.write(ambiente)`).
4. Executa o script/rotina no SAP digitando `script` e seguindo navegação por teclas.
5. Repete tentativa até `max_tentativas=3`.

> Dependência: coordenadas e navegação (ex.: `pyautogui.click(x=637, y=150)`, loops de `down/tab`) precisam bater com a tela real.

### Organização de arquivos
- `separador_historico(base, destino)`
  - Copia o arquivo para `destino` renomeando com a data de modificação (`dd.mm.YYYY`).
- `separador(lista_arquivos, pasta_destino)`
  - Copia arquivos para a pasta destino (cria a pasta se não existir).
- `movimentar_medidores()`, `movimentar_periodicas()`, `movimentar_nf()`, `movimentar_full()`
  - Chamam `separador`/`separador_historico` para diferentes grupos de bases.

### Backup e “substituição” do MRP
- `backup_mrp()`
  - Cria pastas por ano/mês/dia dentro de `config.PASTA_MRP`.
  - Copia `BASE_MRP_PLANEJAMENTO_V2.xlsx` e `MRP_PLAN_V2.xlsm` de `C:\BD\MRP_PLANEJAMENTO` para a pasta do dia.
- `sub_orig()`
  - Copia/atualiza arquivos do MRP para pastas de destino específicas.

### `verificar_bases()`
- Verifica timestamps (`mtime`) de arquivos `TXT_MB52` e `TXT_ME2M`.
- Se não estiverem “no dia esperado”, dispara uma extração no SAP via `extrair_sap` (com script vbs fixo `C:\BD\MRP_v3.vbs`, ambiente `06.01`).

## Agendamento (schedule)
O script inicia uma thread para manter ativo e roda um loop infinito chamando `schedule.run_pending()`.

Horários configurados (principais):
- 07:30 / 07:50 / 08:00
- 08:25 / 08:40
- 11:00 / 11:30
- 13:00 / 13:40
- 13:53 / 14:20
- 15:00 / 15:40
- 17:00 / 17:40
- 18:00 / 19:10 / 19:15 (inclui `verificar_bases`)
- 21:00 / 21:40 / 21:50
- Além disso, no bloco `try`:
  - 00:00: `atualizar_cobertura`
  - 04:00: `atualizar_mrp`
  - 06:00: `backup_mrp`
  - 06:10: `sub_orig`

## Como executar
1. Garanta que a pasta `automacoes/` existe e que `config.py`, `atualizar_mrp.py` e `atualizar_cobertura.py` estão corretos.
2. Instale dependências:
```bash
pip install pyautogui keyboard pygetwindow schedule
```
3. Execute o script principal:
```bash
python seu_script.py
```

> Atenção: pode requerer permissões do Windows para automação (mouse/teclado) e foco em janelas.

## Segurança e boas práticas
- Teste primeiro em um ambiente controlado: `extrair_sap()` depende de coordenadas e títulos.
- Evite rodar junto com outras automações que mexam no SAP.
- Considere logs (arquivo) caso você queira auditoria do que executou em cada horário.

## Licença
Este projeto está **sem licença definida**.

