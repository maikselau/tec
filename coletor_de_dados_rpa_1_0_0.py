from isort import file
import pyautogui
import time

from time import sleep
#from selenium import webdriver
#from webdriver_manager.chrome import ChromeDriverManager
#from selenium.webdriver.chrome.service import Service
#from selenium.webdriver.common.by import By
#from selenium.webdriver.common.keys import Keys
from datetime import datetime
import shutil
import os
import pandas as pd


pasta = '/Users/Administrador/Downloads/'
for f in os.listdir(pasta):
    os.remove(os.path.join(pasta, f))

# abre o filezilla e transfere o arquivo:
pyautogui.hotkey("win", "d")
sleep(2)
pyautogui.press(['l'])
sleep(2)
pyautogui.press(['f'])
pyautogui.press(['enter'])
sleep(3)
pyautogui.hotkey("ctrl", "r")
sleep(10)

pyautogui.press(['tab'])
pyautogui.press(['tab'])
pyautogui.press(['tab'])
pyautogui.press(['tab'])
pyautogui.press(['tab'])
pyautogui.press(['tab'])
pyautogui.press(['tab'])
pyautogui.press(['tab'])
pyautogui.press(['tab'])
pyautogui.press(['tab'])
pyautogui.press(['tab'])
#pyautogui.press(['tab'])
sleep(2)
pyautogui.press(['down'])
pyautogui.press(['down'])
sleep(1)
pyautogui.hotkey("shift", "f10")
pyautogui.press(['down'])
pyautogui.press(['enter'])
sleep(20)
pyautogui.hotkey("alt", "f4")

# extrai o arquivo baixado
pyautogui.hotkey("win", "r")
pyautogui.write('downloads')
pyautogui.press(['enter'])
sleep(2)
pyautogui.press(['up'])
sleep(1)
# pyautogui.press(['enter'])
pyautogui.hotkey("shift", "f10")
pyautogui.press(['down', 'down', 'down', 'down'])
sleep(1)
pyautogui.press(['enter'])
sleep(2)

#nl2 = pd.read_csv("/Users/Administrador/Downloads/ARQ_EXPORTACAO_CHAMADO_FASE.csv", sep=";", on_bad_lines='skip')

#nl2 = pd.read_excel("/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/dataframe.xlsx")


# manipula o arquivo baixado
f_path = "/Users/Administrador/Downloads/ARQ_EXPORTACAO_CHAMADO_FASE.csv"
t = os.path.getctime(f_path)

t_str = time.ctime(t)
t_obj = time.strptime(t_str)
form_t = time.strftime("%Y-%m-%d %H:%M:%S", t_obj)


form_t = form_t.replace(":", "꞉")
os.rename(
    f_path, os.path.split(f_path)[0] + '/' + form_t + os.path.splitext(f_path)[1])

# move o arquivo
nome_doc = "/Users/Administrador/Downloads/" + \
    form_t + os.path.splitext(f_path)[1]
nome_doc2 = "/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/base_dados/" + \
    form_t + os.path.splitext(f_path)[1]

shutil.move(
    nome_doc, nome_doc2)

novo_nome_doc = "/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/base_dados/" + \
    form_t + os.path.splitext(f_path)[1]

# comeca a montagem do dataframe

# carrega o arquivo do dataframe

print("Começando o carregamento do arquivo do dataframe")


df = pd.read_excel("/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/dataframe.xlsx")

# carrega o arquivo da nova leitura
nl = pd.read_csv(novo_nome_doc, sep=";", on_bad_lines='skip')


# exclui linhas e colunas vazias
nl = nl.dropna(how='all')
nl = nl.dropna(how='all', axis=1)


# concatena o DF com o NL
df = pd.concat([df, nl])

print("Concatenagem do dataframe concluida")

df['CHAMADO'] = df['CHAMADO'].replace(',0', '', regex=True)
df['CASO'] = df['CASO'].replace(',0', '', regex=True)
df['CONTADOR'] = df['CONTADOR'].replace(',0', '', regex=True)

# sobrescreve o frame

df.to_excel("/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/dataframe.xlsx", index=False)

# chama o frame atualizado
df = pd.read_excel(
    "/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/dataframe.xlsx")

df = df.drop_duplicates(subset=['CASO'], keep='last')

# sobrescreve o frame
df.to_excel("/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/dataframe.xlsx", index=False)


print("Terminada a montagem do dataframe, agora vou manipular os dados")


# faz uma serie de substituições devido aos nomes dos clientes, gera as colunas de cidade, estado e cliente:
df['LOCALIZACAO'] = df['LOCALIZACAO'].str.replace(
    '- ITAJAI SHOPPING', '', regex=True)
df['LOCALIZACAO'] = df['LOCALIZACAO'].str.replace(
    'RIOXEL - RS', 'RIOXEL - GUAIBA - RS', regex=True)
divcliente = df['CLIENTE'].str.split('-')
divisao = df['LOCALIZACAO'].str.split(' - ')
cidade = divisao.str.get(1)
Estado = divisao.str.get(2)
cliente = divcliente.str.get(0)
df['CIDADE'] = cidade
df['ESTADO'] = Estado
df['NOME CLIENTE'] = cliente
df['ESTADO'] = df['ESTADO'].str.replace('SANTA CATARINA', 'SC', regex=True)
df['ESTADO'] = df['ESTADO'].str.replace('EAD', 'SC', regex=True)
df['ESTADO'] = df['ESTADO'].str.replace('FACULDADE', 'SC', regex=True)
df['ESTADO'] = df['ESTADO'].str.replace('BLUMENGARTEN', 'SC', regex=True)
df['ESTADO'] = df['ESTADO'].str.replace('ITAJAI', 'SC', regex=True)
df['CIDADE'] = df['CIDADE'].str.replace('SSP IGP', 'ITAJAI', regex=True)
df['CIDADE'] = df['CIDADE'].str.replace('PR', 'CURITIBA', regex=True)
pr = 'PR'
df['ESTADO'] = df['ESTADO'].fillna(pr)

# Bloco de manipulação de chamados internos:
# troca por NaN o CPF dos residentes para excluir dos relatorios chamados internos abertos pelo campo:
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '03843468044', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '103.312.189-41', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '045.930.750-92', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '02599073910', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '00802384080', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '36877972015', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '00000001', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '50888323972', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '078.369.129-73', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '02817648080', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '94921555087', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '83017089053', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '07859662930', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '03500208924', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '10264900901', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '046.640.721-12', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '08573363983', '', regex=True)
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].replace(
    '085.733.639-83', '', regex=True)
df['RELATOR_MATRICULA'] = pd.to_numeric(
    df['RELATOR_MATRICULA'], errors="coerce")
df['RELATOR_MATRICULA'] = df['RELATOR_MATRICULA'].fillna('1')

# descobre o indice das matriculas que precisam dropar e dropa:
indice_matricula = df[(df['RELATOR_MATRICULA'] != '1')].index
df.drop(indice_matricula, inplace=True)
df['RELATOR_MATRICULA'] = pd.to_numeric(
    df['RELATOR_MATRICULA'], errors="coerce")

# dropa cliente CHAT-ROBO:
indice_cliente_chatrobo = df[(df['CLIENTE'] == 'CHAT-ROBO')].index
df.drop(indice_cliente_chatrobo, inplace=True)

# dropa cliente TECPRINTERS:
indice_cliente_tecprinters = df[(df['CLIENTE'] == 'TECPRINTERS')].index
df.drop(indice_cliente_tecprinters, inplace=True)

# dropa chamados sem o bem:
indice_bem = df[(df['BEM'] == '')].index
df.drop(indice_bem, inplace=True)

# dropa tipo Exemplos:
indice_tipo_exemplos = df[(df['TIPO'] == 'Exemplos')].index
df.drop(indice_tipo_exemplos, inplace=True)

# dropa chamado sem caso:
indice_semcaso = df[(df['CASO'] == '')].index
df.drop(indice_semcaso, inplace=True)

# dropa localizacao a definir:
indice_localizacao = df[(df['LOCALIZACAO'] == 'a definir')].index
df.drop(indice_localizacao, inplace=True)

# dropa status "aberto":
indice_aberto = df[(df['STATUS'] == 'ABERTO')].index
df.drop(indice_aberto, inplace=True)

# ajusta nome de cliente errado:
df['NOME CLIENTE'] = df['NOME CLIENTE'].str.replace(
    'COMPANHIA SANEPAR (PROVISORIO)', 'SANEPAR', regex=True)
df['NOME CLIENTE'] = df['NOME CLIENTE'].str.replace(
    'PM XANGRI', 'PM XANGRI-LA', regex=True)

# dropa contextos que não forem 1:
indice_contexto = df[(df['CONTEXTO'] == '2')].index
df.drop(indice_contexto, inplace=True)



print("Terminei de dropar coisas erradas")

# ajusta os tipos selecionados errados na abertura do chamado:
# Bloco de incidentes de digitalização
df['TIPO'] = df['TIPO'].str.replace(
    'SCANNER - ATOLAMENTO DE PAPEL NO ADF', 'Impressora - ID - Atolamento de papel no ADF', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'SCANNER - CONFIGURAÇÃO DE DIGITALIZAÇÃO', 'Impressora - ID - Configuração de Digitalização', regex=True)
df['TIPO'] = df['TIPO'].str.replace('SCANNER - E-MAIL DIGITALIZADO NÃO É RECEBIDO',
                                    'Impressora - ID - E-mail digitalizado não é recebido', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'SCANNER - ERRO DE CONEXÃO COM O DESTINO', 'Impressora - ID - Erro de Conexão com Destino', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'SCANNER - FALHAS, RISCOS OU MANCHAS', 'Impressora - ID - Manchas ou Falhas no arquivo', regex=True)
df['TIPO'] = df['TIPO'].str.replace('SCANNER - ATIVAR OU DESATIVAR FRENTE E VERSO',
                                    'Impressora - ID - Não ativa/desativa Frente e Verso', regex=True)
df['TIPO'] = df['TIPO'].str.replace('SCANNER - ATIVAR OU DESATIVAR FRENTE E VERSO',
                                    'Impressora - ID - Não ativa/desativa Frente e Verso', regex=True)
# Bloco de incidentes de hardware
df['TIPO'] = df['TIPO'].str.replace(
    'HARDWARE - CÓDIGO DE ERRO NA TELA', 'Impressora - Inc. HW - Apresenta Erro na Tela', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'HARDWARE - TELA TRAVADA', 'Impressora - Inc. HW - Apresenta Travamento da Tela', regex=True)
df['TIPO'] = df['TIPO'].str.replace('HARDWARE - IMPRESSORA TRAVADA',
                                    'Impressora - Inc. HW - Apresenta Travamento Total do Equipamento', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'HARDWARE - ATOLAMENTO DE PAPEL', 'Impressora - Inc. HW - Atolamento de Papel', regex=True)
df['TIPO'] = df['TIPO'].str.replace('HARDWARE - CONVERTER COLORIDO PARA MONOCROMATICO',
                                    'Impressora - Inc. HW - Conversão de Colorido para Monocromático', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'HARDWARE - SENSOR COM ERRO', 'Impressora - Inc. HW - Defeito em Sensores', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'HARDWARE - IMPRESSORA NÃO LIGA', 'Impressora - Inc. HW - Equipamento não Liga', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'HARDWARE - PEÇA DANIFICADA OU QUEBRADA', 'Impressora - Inc. HW - Peça Danificada/Quebrada', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'HARDWARE - TRANSFORMADOR NÃO LIGA', 'Impressora - Inc. HW - Transformador não Liga', regex=True)
# Bloco de incidentes de impressão
df['TIPO'] = df['TIPO'].str.replace(
    'IMPRESSÃO - ERRO AO IMPRIMIR', 'Impressora - Inc. Imp. - Erro Fila de Impressão', regex=True)
df['TIPO'] = df['TIPO'].str.replace('IMPRESSÃO - MANCHAS, RISCOS OU FALHAS',
                                    'Impressora - Inc. Imp. - Impressões Manchadas ou com Falhas', regex=True)
df['TIPO'] = df['TIPO'].str.replace('IMPRESSÃO - CARACTERES ESTRANHOS',
                                    'Impressora - Inc. Imp. - Imprimindo Caracteres Estranhos', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'IMPRESSÃO - LENTIDÃO AO IMPRIMIR', 'Impressora - Inc. Imp. - Lentidão para Imprimir', regex=True)
df['TIPO'] = df['TIPO'].str.replace('IMPRESSÃO - ATIVAR/DESATIVAR FRENTE E VERSO',
                                    'Impressora - Inc. Imp. - Não ativa/desativa Frente e Verso', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'IMPRESSÃO - NÃO IMPRIME', 'Impressora - Inc. Imp. - Não Imprime', regex=True)
df['TIPO'] = df['TIPO'].str.replace('IMPRESSÃO - NÃO ENCONTREI A OPÇÃO DESEJADA',
                                    'Impressora - Opção de chamado não encontrada no SB3', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'IMPRESSÃO - SUGESTÃO DE MELHORIA', 'Impressora - Sugestão de Melhoria', regex=True)
# Bloco de incidentes de suprimentos
df['TIPO'] = df['TIPO'].str.replace(
    'SUPRIMENTOS - FALTA DE TONER', 'Impressora - Inc. Sup. - Falta de Toner', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'SUPRIMENTOS - FOTOCONDUTOR', 'Impressora - Inc. Sup. - Fotocondutor', regex=True)
df['TIPO'] = df['TIPO'].str.replace('SUPRIMENTOS - RECOLHIMENTO DE CARCAÇAS',
                                    'Impressora - Inc. Sup. - Recolhimento de carcaças', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'SUPRIMENTOS - TONER VAZANDO', 'Impressora - Inc. Sup. - Toner com Vazamento', regex=True)
df['TIPO'] = df['TIPO'].str.replace('SUPRIMENTOS - TONER NÃO RECONHECIDO',
                                    'Impressora - Inc. Sup. - Toner não foi Reconhecido', regex=True)
# Bloco de RDS
df['TIPO'] = df['TIPO'].str.replace('REQUISICAO DE SERVIÇO - ACESSO AO RELATORIO DE CONTADORES',
                                    'Impressora - RDS - Acesso ao relatório de contadores', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'REQUISICAO DE SERVIÇO - EXCLUSÃO DE USUÁRIO NDD', 'Impressora - RDS - Exclusão de usuário NDD', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - INSTALAÇÃO DE IMPRESSORA', 'Impressora - RDS - Instalação de Impressora', regex=True)
df['TIPO'] = df['TIPO'].str.replace('REQUISIÇÃO DE SERVIÇO - INSTALAÇÃO DE SOFTWARE DE MONITORAMENTO',
                                    'Impressora - RDS - Instalação de Software de Monitoramento', regex=True)
df['TIPO'] = df['TIPO'].str.replace('REQUISIÇÃO DE SERVIÇO - MOVIMENTAÇÃO DE IMPRESSORA',
                                    'Impressora - RDS - Movimentação/Remanejo de Impressora', regex=True)
df['TIPO'] = df['TIPO'].str.replace('REQUISIÇÃO DE SERVIÇO - RECOLHIMENTO DE IMPRESSORA',
                                    'Impressora - RDS - Recolhimento de Impressora', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - SUPRIMENTO RESERVA', 'Impressora - RDS - Suprimento Reserva', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - TROCA DE REDE FÍSICA', 'Impressora - RDS - Troca de Rede Fisica', regex=True)
df['TIPO'] = df['TIPO'].str.replace(
    'REQUISIÇÕES DE SERVIÇO - TROCA DE REDE LÓGICA', 'Impressora - RDS - Troca de Rede Lógica', regex=True)


# ajusta os tipos selecionados errados na abertura do caso:
# Bloco de incidentes de digitalização
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SCANNER - ATOLAMENTO DE PAPEL NO ADF', 'Impressora - ID - Atolamento de papel no ADF', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SCANNER - CONFIGURAÇÃO DE DIGITALIZAÇÃO', 'Impressora - ID - Configuração de Digitalização', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SCANNER - E-MAIL DIGITALIZADO NÃO É RECEBIDO', 'Impressora - ID - E-mail digitalizado não é recebido', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SCANNER - ERRO DE CONEXÃO COM O DESTINO', 'Impressora - ID - Erro de Conexão com Destino', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SCANNER - FALHAS, RISCOS OU MANCHAS', 'Impressora - ID - Manchas ou Falhas no arquivo', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SCANNER - ATIVAR OU DESATIVAR FRENTE E VERSO', 'Impressora - ID - Não ativa/desativa Frente e Verso', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SCANNER - ATIVAR OU DESATIVAR FRENTE E VERSO', 'Impressora - ID - Não ativa/desativa Frente e Verso', regex=True)
# Bloco de incidentes de hardware
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'HARDWARE - CÓDIGO DE ERRO NA TELA', 'Impressora - Inc. HW - Apresenta Erro na Tela', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'HARDWARE - TELA TRAVADA', 'Impressora - Inc. HW - Apresenta Travamento da Tela', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'HARDWARE - IMPRESSORA TRAVADA', 'Impressora - Inc. HW - Apresenta Travamento Total do Equipamento', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'HARDWARE - ATOLAMENTO DE PAPEL', 'Impressora - Inc. HW - Atolamento de Papel', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'HARDWARE - CONVERTER COLORIDO PARA MONOCROMATICO', 'Impressora - Inc. HW - Conversão de Colorido para Monocromático', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'HARDWARE - SENSOR COM ERRO', 'Impressora - Inc. HW - Defeito em Sensores', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'HARDWARE - IMPRESSORA NÃO LIGA', 'Impressora - Inc. HW - Equipamento não Liga', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'HARDWARE - PEÇA DANIFICADA OU QUEBRADA', 'Impressora - Inc. HW - Peça Danificada/Quebrada', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'HARDWARE - TRANSFORMADOR NÃO LIGA', 'Impressora - Inc. HW - Transformador não Liga', regex=True)
# Bloco de incidentes de impressão
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'IMPRESSÃO - ERRO AO IMPRIMIR', 'Impressora - Inc. Imp. - Erro Fila de Impressão', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'IMPRESSÃO - MANCHAS, RISCOS OU FALHAS', 'Impressora - Inc. Imp. - Impressões Manchadas ou com Falhas', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'IMPRESSÃO - CARACTERES ESTRANHOS', 'Impressora - Inc. Imp. - Imprimindo Caracteres Estranhos', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'IMPRESSÃO - LENTIDÃO AO IMPRIMIR', 'Impressora - Inc. Imp. - Lentidão para Imprimir', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'IMPRESSÃO - ATIVAR/DESATIVAR FRENTE E VERSO', 'Impressora - Inc. Imp. - Não ativa/desativa Frente e Verso', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'IMPRESSÃO - NÃO IMPRIME', 'Impressora - Inc. Imp. - Não Imprime', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'IMPRESSÃO - NÃO ENCONTREI A OPÇÃO DESEJADA', 'Impressora - Opção de chamado não encontrada no SB3', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'IMPRESSÃO - SUGESTÃO DE MELHORIA', 'Impressora - Sugestão de Melhoria', regex=True)
# Bloco de incidentes de suprimentos
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SUPRIMENTOS - FALTA DE TONER', 'Impressora - Inc. Sup. - Falta de Toner', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SUPRIMENTOS - FOTOCONDUTOR', 'Impressora - Inc. Sup. - Fotocondutor', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SUPRIMENTOS - RECOLHIMENTO DE CARCAÇAS', 'Impressora - Inc. Sup. - Recolhimento de carcaças', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SUPRIMENTOS - TONER VAZANDO', 'Impressora - Inc. Sup. - Toner com Vazamento', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'SUPRIMENTOS - TONER NÃO RECONHECIDO', 'Impressora - Inc. Sup. - Toner não foi Reconhecido', regex=True)
# Bloco de RDS
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'REQUISICAO DE SERVIÇO - ACESSO AO RELATORIO DE CONTADORES', 'Impressora - RDS - Acesso ao relatório de contadores', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'REQUISICAO DE SERVIÇO - EXCLUSÃO DE USUÁRIO NDD', 'Impressora - RDS - Exclusão de usuário NDD', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - INSTALAÇÃO DE IMPRESSORA', 'Impressora - RDS - Instalação de Impressora', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - INSTALAÇÃO DE SOFTWARE DE MONITORAMENTO', 'Impressora - RDS - Instalação de Software de Monitoramento', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - MOVIMENTAÇÃO DE IMPRESSORA', 'Impressora - RDS - Movimentação/Remanejo de Impressora', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - RECOLHIMENTO DE IMPRESSORA', 'Impressora - RDS - Recolhimento de Impressora', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - SUPRIMENTO RESERVA', 'Impressora - RDS - Suprimento Reserva', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - TROCA DE REDE FÍSICA', 'Impressora - RDS - Troca de Rede Fisica', regex=True)
df['TIPO_CASO'] = df['TIPO_CASO'].str.replace(
    'REQUISIÇÕES DE SERVIÇO - TROCA DE REDE LÓGICA', 'Impressora - RDS - Troca de Rede Lógica', regex=True)


# ajusta os tipos selecionados errados na abertura do caso anterior:
# Bloco de incidentes de digitalização
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].astype(str).str.replace('SCANNER - ATOLAMENTO DE PAPEL NO ADF', 'Impressora - ID - Atolamento de papel no ADF', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SCANNER - CONFIGURAÇÃO DE DIGITALIZAÇÃO', 'Impressora - ID - Configuração de Digitalização', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SCANNER - E-MAIL DIGITALIZADO NÃO É RECEBIDO', 'Impressora - ID - E-mail digitalizado não é recebido', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SCANNER - ERRO DE CONEXÃO COM O DESTINO', 'Impressora - ID - Erro de Conexão com Destino', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SCANNER - FALHAS, RISCOS OU MANCHAS', 'Impressora - ID - Manchas ou Falhas no arquivo', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SCANNER - ATIVAR OU DESATIVAR FRENTE E VERSO', 'Impressora - ID - Não ativa/desativa Frente e Verso', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SCANNER - ATIVAR OU DESATIVAR FRENTE E VERSO', 'Impressora - ID - Não ativa/desativa Frente e Verso', regex=True)
# Bloco de incidentes de hardware
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'HARDWARE - CÓDIGO DE ERRO NA TELA', 'Impressora - Inc. HW - Apresenta Erro na Tela', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'HARDWARE - TELA TRAVADA', 'Impressora - Inc. HW - Apresenta Travamento da Tela', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'HARDWARE - IMPRESSORA TRAVADA', 'Impressora - Inc. HW - Apresenta Travamento Total do Equipamento', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'HARDWARE - ATOLAMENTO DE PAPEL', 'Impressora - Inc. HW - Atolamento de Papel', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'HARDWARE - CONVERTER COLORIDO PARA MONOCROMATICO', 'Impressora - Inc. HW - Conversão de Colorido para Monocromático', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'HARDWARE - SENSOR COM ERRO', 'Impressora - Inc. HW - Defeito em Sensores', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'HARDWARE - IMPRESSORA NÃO LIGA', 'Impressora - Inc. HW - Equipamento não Liga', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'HARDWARE - PEÇA DANIFICADA OU QUEBRADA', 'Impressora - Inc. HW - Peça Danificada/Quebrada', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'HARDWARE - TRANSFORMADOR NÃO LIGA', 'Impressora - Inc. HW - Transformador não Liga', regex=True)
# Bloco de incidentes de impressão
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'IMPRESSÃO - ERRO AO IMPRIMIR', 'Impressora - Inc. Imp. - Erro Fila de Impressão', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'IMPRESSÃO - MANCHAS, RISCOS OU FALHAS', 'Impressora - Inc. Imp. - Impressões Manchadas ou com Falhas', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'IMPRESSÃO - CARACTERES ESTRANHOS', 'Impressora - Inc. Imp. - Imprimindo Caracteres Estranhos', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'IMPRESSÃO - LENTIDÃO AO IMPRIMIR', 'Impressora - Inc. Imp. - Lentidão para Imprimir', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'IMPRESSÃO - ATIVAR/DESATIVAR FRENTE E VERSO', 'Impressora - Inc. Imp. - Não ativa/desativa Frente e Verso', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'IMPRESSÃO - NÃO IMPRIME', 'Impressora - Inc. Imp. - Não Imprime', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'IMPRESSÃO - NÃO ENCONTREI A OPÇÃO DESEJADA', 'Impressora - Opção de chamado não encontrada no SB3', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'IMPRESSÃO - SUGESTÃO DE MELHORIA', 'Impressora - Sugestão de Melhoria', regex=True)
# Bloco de incidentes de suprimentos
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SUPRIMENTOS - FALTA DE TONER', 'Impressora - Inc. Sup. - Falta de Toner', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SUPRIMENTOS - FOTOCONDUTOR', 'Impressora - Inc. Sup. - Fotocondutor', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SUPRIMENTOS - RECOLHIMENTO DE CARCAÇAS', 'Impressora - Inc. Sup. - Recolhimento de carcaças', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SUPRIMENTOS - TONER VAZANDO', 'Impressora - Inc. Sup. - Toner com Vazamento', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'SUPRIMENTOS - TONER NÃO RECONHECIDO', 'Impressora - Inc. Sup. - Toner não foi Reconhecido', regex=True)
# Bloco de RDS
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'REQUISICAO DE SERVIÇO - ACESSO AO RELATORIO DE CONTADORES', 'Impressora - RDS - Acesso ao relatório de contadores', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'REQUISICAO DE SERVIÇO - EXCLUSÃO DE USUÁRIO NDD', 'Impressora - RDS - Exclusão de usuário NDD', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - INSTALAÇÃO DE IMPRESSORA', 'Impressora - RDS - Instalação de Impressora', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - INSTALAÇÃO DE SOFTWARE DE MONITORAMENTO', 'Impressora - RDS - Instalação de Software de Monitoramento', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - MOVIMENTAÇÃO DE IMPRESSORA', 'Impressora - RDS - Movimentação/Remanejo de Impressora', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - RECOLHIMENTO DE IMPRESSORA', 'Impressora - RDS - Recolhimento de Impressora', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - SUPRIMENTO RESERVA', 'Impressora - RDS - Suprimento Reserva', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'REQUISIÇÃO DE SERVIÇO - TROCA DE REDE FÍSICA', 'Impressora - RDS - Troca de Rede Fisica', regex=True)
df['TIPO_CASO_ANTERIOR'] = df['TIPO_CASO_ANTERIOR'].str.replace(
    'REQUISIÇÕES DE SERVIÇO - TROCA DE REDE LÓGICA', 'Impressora - RDS - Troca de Rede Lógica', regex=True)

# Ajuste dos nomes das colunas:
df.rename(
    columns={'DT_ABERTURA': 'DATA E HORA DA ABERTURA DO CHAMADO'}, inplace=True)
df.rename(
    columns={'DT_LIMITE': 'DATA E HORA LIMITE PARA ATENDIMENTO'}, inplace=True)
df.rename(columns={
          'DT_ATUALIZACAO_CASO': 'DATA E HORA DA ULTIMA ATUALIZAÇÃO NO CASO'}, inplace=True)
df.rename(columns={
          'DT_ATUALIZACAO_CHAMADO': 'DATA E HORA DO FECHAMENTO DO CHAMADO'}, inplace=True)
df.rename(columns={'CLIENTE': 'SITE'}, inplace=True)
df.rename(
    columns={'INF_ADIC_CLIENTE': 'INFORMAÇÃO ADICIONAL SOBRE O SITE'}, inplace=True)
df.rename(columns={'PRIORIDADE': 'TEMPO DE SLA'}, inplace=True)
df.rename(columns={'STATUS': 'STATUS DO CHAMADO'}, inplace=True)
df.rename(columns={'SLA': 'STATUS DO SLA'}, inplace=True)
df.rename(columns={'TIPO': 'CATEGORIA QUE O CHAMADO FOI ABERTO'}, inplace=True)
df.rename(
    columns={'CLASS_ENCERRAMENTO': 'CLASSIFICAÇÃO NO ENCERRAMENTO'}, inplace=True)
df.rename(columns={'FASE_CASO': 'FASE NO ENCERRAMENTO'}, inplace=True)
df.rename(columns={'DESCRICAO': 'DESCRIÇÃO DO CHAMADO'}, inplace=True)
df.rename(columns={'SOLUCAO': 'SOLUÇÃO DO CHAMADO'}, inplace=True)
df.rename(columns={'RESPONSAVEL': 'RESPONSÁVEL ATUAL'}, inplace=True)
df.rename(
    columns={'GRUPO_RESPONSAVEL': 'GRUPO DO RESPONSÁVEL ATUAL'}, inplace=True)
df.rename(columns={
          'NR_TEMPO_DURACAO_CHAMADO': 'TEMPO DE DURAÇÃO DO CHAMADO'}, inplace=True)
df.rename(
    columns={'NR_ATRASO_CHAMADO': 'TEMPO DE ATRASO NO ATENDIMENTO'}, inplace=True)
df.rename(columns={
          'NR_ATRASO_DESC_INTERR_CASO': 'TEMPO DE ATRASO MENOS A INTERRUPÇÃO'}, inplace=True)
df.rename(columns={
          'RESPONSAVEL_ENCERR': 'RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'}, inplace=True)
df.rename(
    columns={'TIPO_CASO': 'CATEGORIA QUE O CHAMADO FOI ENCERRADO'}, inplace=True)
df.rename(columns={'TIPO_CASO_ANTERIOR': 'CATEGORIA ANTERIOR'}, inplace=True)
df.rename(columns={'FASE1': 'FASE 1 DO ATENDIMENTO'}, inplace=True)
df.rename(
    columns={'ESPECIALISTA_FASE1': 'ESPECIALISTA QUE ATUOU NA FASE 1'}, inplace=True)
df.rename(
    columns={'DT_Fase1_INI': 'DATA E HORA QUE ENTROU NA FASE 1'}, inplace=True)
df.rename(
    columns={'DT_Fase1_OUT': 'DATA E HORA QUE SAIU DA FASE 1'}, inplace=True)
df.rename(
    columns={'DT_Fase1_OUT': 'DATA E HORA QUE SAIU DA FASE 1'}, inplace=True)
df.rename(columns={'SLA_FASE1': 'STATUS DO SLA NA FASE 1'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE1': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 1'}, inplace=True)
df.rename(columns={'FASE2': 'FASE 2 DO ATENDIMENTO'}, inplace=True)
df.rename(
    columns={'ESPECIALISTA_FASE2': 'ESPECIALISTA QUE ATUOU NA FASE 2'}, inplace=True)
df.rename(
    columns={'DT_Fase2_INI': 'DATA E HORA QUE ENTROU NA FASE 2'}, inplace=True)
df.rename(
    columns={'DT_Fase2_OUT': 'DATA E HORA QUE SAIU DA FASE 2'}, inplace=True)
df.rename(
    columns={'DT_Fase2_OUT': 'DATA E HORA QUE SAIU DA FASE 2'}, inplace=True)
df.rename(columns={'SLA_FASE2': 'STATUS DO SLA NA FASE 2'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE2': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 2'}, inplace=True)
df.rename(columns={'FASE3': 'FASE 3 DO ATENDIMENTO'}, inplace=True)
df.rename(
    columns={'ESPECIALISTA_FASE3': 'ESPECIALISTA QUE ATUOU NA FASE 3'}, inplace=True)
df.rename(
    columns={'DT_Fase3_INI': 'DATA E HORA QUE ENTROU NA FASE 3'}, inplace=True)
df.rename(
    columns={'DT_Fase3_OUT': 'DATA E HORA QUE SAIU DA FASE 3'}, inplace=True)
df.rename(
    columns={'DT_Fase3_OUT': 'DATA E HORA QUE SAIU DA FASE 3'}, inplace=True)
df.rename(columns={'SLA_FASE3': 'STATUS DO SLA NA FASE 3'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE3': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 3'}, inplace=True)
df.rename(columns={'FASE4': 'FASE 4 DO ATENDIMENTO'}, inplace=True)
df.rename(
    columns={'ESPECIALISTA_FASE4': 'ESPECIALISTA QUE ATUOU NA FASE 4'}, inplace=True)
df.rename(
    columns={'DT_Fase4_INI': 'DATA E HORA QUE ENTROU NA FASE 4'}, inplace=True)
df.rename(
    columns={'DT_Fase4_OUT': 'DATA E HORA QUE SAIU DA FASE 4'}, inplace=True)
df.rename(
    columns={'DT_Fase4_OUT': 'DATA E HORA QUE SAIU DA FASE 4'}, inplace=True)
df.rename(columns={'SLA_FASE4': 'STATUS DO SLA NA FASE 4'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE4': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 4'}, inplace=True)
df.rename(columns={'FASE5': 'FASE 5 DO ATENDIMENTO'}, inplace=True)
df.rename(
    columns={'ESPECIALISTA_FASE5': 'ESPECIALISTA QUE ATUOU NA FASE 5'}, inplace=True)
df.rename(
    columns={'DT_Fase5_INI': 'DATA E HORA QUE ENTROU NA FASE 5'}, inplace=True)
df.rename(
    columns={'DT_Fase5_OUT': 'DATA E HORA QUE SAIU DA FASE 5'}, inplace=True)
df.rename(
    columns={'DT_Fase5_OUT': 'DATA E HORA QUE SAIU DA FASE 5'}, inplace=True)
df.rename(columns={'SLA_FASE5': 'STATUS DO SLA NA FASE 5'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE5': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 5'}, inplace=True)
df.rename(columns={'FASE6': 'FASE 6 DO ATENDIMENTO'}, inplace=True)
df.rename(
    columns={'ESPECIALISTA_FASE6': 'ESPECIALISTA QUE ATUOU NA FASE 6'}, inplace=True)
df.rename(
    columns={'DT_Fase6_INI': 'DATA E HORA QUE ENTROU NA FASE 6'}, inplace=True)
df.rename(
    columns={'DT_Fase6_OUT': 'DATA E HORA QUE SAIU DA FASE 6'}, inplace=True)
df.rename(
    columns={'DT_Fase6_OUT': 'DATA E HORA QUE SAIU DA FASE 6'}, inplace=True)
df.rename(columns={'SLA_FASE6': 'STATUS DO SLA NA FASE 6'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE6': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 6'}, inplace=True)
df.rename(columns={'FASE7': 'FASE 7 DO ATENDIMENTO'}, inplace=True)
df.rename(
    columns={'ESPECIALISTA_FASE7': 'ESPECIALISTA QUE ATUOU NA FASE 7'}, inplace=True)
df.rename(
    columns={'DT_Fase7_INI': 'DATA E HORA QUE ENTROU NA FASE 7'}, inplace=True)
df.rename(
    columns={'DT_Fase7_OUT': 'DATA E HORA QUE SAIU DA FASE 7'}, inplace=True)
df.rename(
    columns={'DT_Fase7_OUT': 'DATA E HORA QUE SAIU DA FASE 7'}, inplace=True)
df.rename(columns={'SLA_FASE7': 'STATUS DO SLA NA FASE 7'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE7': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 7'}, inplace=True)
df.rename(columns={'FASE8': 'FASE 8 DO ATENDIMENTO'}, inplace=True)
df.rename(
    columns={'ESPECIALISTA_FASE8': 'ESPECIALISTA QUE ATUOU NA FASE 8'}, inplace=True)
df.rename(
    columns={'DT_Fase8_INI': 'DATA E HORA QUE ENTROU NA FASE 8'}, inplace=True)
df.rename(
    columns={'DT_Fase8_OUT': 'DATA E HORA QUE SAIU DA FASE 8'}, inplace=True)
df.rename(
    columns={'DT_Fase8_OUT': 'DATA E HORA QUE SAIU DA FASE 8'}, inplace=True)
df.rename(columns={'SLA_FASE8': 'STATUS DO SLA NA FASE 8'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE8': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 8'}, inplace=True)
df.rename(columns={'FASE9': 'FASE 9 DO ATENDIMENTO'}, inplace=True)
df.rename(
    columns={'ESPECIALISTA_FASE9': 'ESPECIALISTA QUE ATUOU NA FASE 9'}, inplace=True)
df.rename(
    columns={'DT_Fase9_INI': 'DATA E HORA QUE ENTROU NA FASE 9'}, inplace=True)
df.rename(
    columns={'DT_Fase9_OUT': 'DATA E HORA QUE SAIU DA FASE 9'}, inplace=True)
df.rename(
    columns={'DT_Fase9_OUT': 'DATA E HORA QUE SAIU DA FASE 9'}, inplace=True)
df.rename(columns={'SLA_FASE9': 'STATUS DO SLA NA FASE 9'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE9': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 9'}, inplace=True)
df.rename(columns={'FASE10': 'FASE 10 DO ATENDIMENTO'}, inplace=True)
df.rename(columns={
          'ESPECIALISTA_FASE10': 'ESPECIALISTA QUE ATUOU NA FASE 10'}, inplace=True)
df.rename(
    columns={'DT_Fase10_INI': 'DATA E HORA QUE ENTROU NA FASE 10'}, inplace=True)
df.rename(
    columns={'DT_Fase10_OUT': 'DATA E HORA QUE SAIU DA FASE 10'}, inplace=True)
df.rename(
    columns={'DT_Fase10_OUT': 'DATA E HORA QUE SAIU DA FASE 10'}, inplace=True)
df.rename(columns={'SLA_FASE10': 'STATUS DO SLA NA FASE 10'}, inplace=True)
df.rename(columns={
          'ACOMPANHAMENTO_FASE10': 'ACOMPANHAMENTO INSERIDO PELO ESPECIALISTA NA FASE 10'}, inplace=True)

df.rename(columns={'TIPO_BEM': 'MODELO DO EQUIPAMENTO'}, inplace=True)          

# Ajuste dos nomes dos especialistas:
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'ewerton.lara', 'EWERTON LARA', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'evandro.duarte', 'EVANDRO DUARTE', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'alexandre.medeiros', 'ALEXANDRE FERNANDES MEDEIROS', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'allan.souza', 'ALLAN DE SOUZA QUADRA', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'augusto.alves', 'AUGUSTO PEREIRA ALVES', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'clairton.lima', 'CLAIRTON ALVES DE LIMA', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'claudio.valenza', 'CLAUDIO VALENZA JUNIOR', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'daniel', 'DANIEL DIETRICH OLIVEIRA', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'denison.pereira', 'DENISON RODRIGO', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'diego.neimayer', 'DIEGO JOSE RISE NEIMAYER', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'djair.bento', 'DJAIR JOSE BENTO JUNIOR', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'gabriel.ayres', 'GABRIEL FAÇANHA LOPES AYRES', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'gabriel.rudiniski', 'GABRIEL FRANCISCO LORENCO RUDINISKI', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'gilvando', 'GILVANDO CARPINELLI', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'kleber', 'KLEBER SALOMAO HACK', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'Usuário para Abertura de Chamados via CHAT-ROBO', '', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'allan', 'ALLAN JONES DOS SANTOS MEDEIROS MACIEL', regex=True)
df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'] = df['RESPONSÁVEL PELO ENCERRAMENTO DO CHAMADO'].str.replace(
    'alvaro.laurentino', 'ALVARO LAURENTINO', regex=True)


print("terminei de ajustar os nomes e vou comecar a atualizar os relatorios e salvar")

#Cria coluna de data de atualização do banco de dados:

df['DATA E HORA DA ULTIMA ATUALIZACAO DO BANCO DE DADOS'] = time.strftime('%d/%m/%y %H:%M:%S', time.localtime())

#salva as versoes do relatorio para cada setor:

df.to_excel("/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/relatorio_base_tratada.xlsx", index=False)

df = pd.read_excel("/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/relatorio_base_tratada.xlsx")

df['CATEGORIA ANTERIOR'] = df['CATEGORIA ANTERIOR'].str.replace('nan', '', regex=True)

print("Terminei a base tratada")

df.to_excel("/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/relatorio_base_bi.xlsx", index=False)

df = df.drop(columns=['DATA E HORA DA ULTIMA ATUALIZACAO DO BANCO DE DADOS'])

df.to_excel("/Users/Administrador/TECPRINTERS TECNOLOGIA DE IMPRESSAO LTDA/Portal Tecprinters - RPA/Criador de Dashboards/relatorio_base_tratada.xlsx", index=False)

print("Terminei o base BI")

print("Tchau, obrigado!")