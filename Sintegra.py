import time
import pygsheets #  biblioteca para mexer com o google sheets
import logging  # gerador de logs
from urllib.request import urlopen
import json

#**************** Inicio ****************

#  local do token de autorização da API do Google Sheets
path = ('Credenciais.json')
gc = pygsheets.authorize(service_account_file=path)

# abrindo a planilha, e escolhendo a aba
planilha = gc.open_by_key('id da planilha')
aba = planilha.worksheet_by_title('Sugestão_Procura')  # sh[0] seleciona a primeira sheet

# log config básico. 
logging.basicConfig(filename='Consulta_Sintegra.log', filemode='a', format='%(asctime)s - %(levelname)s - %(message)s')

#informando no log que essa etapa foi completada
logging.warning('Conexão com a planilha bem sucedida.')
print('Conexão com a planilha bem sucedida.')

# Calcular quantidade de pesquisas
celulas = aba.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix')
total_celulas = len(celulas)

total = str(total_celulas)
i = 2

print('Total de consultas a realizar: ' + total)

while i <= total_celulas:

    try:
        cnpj2 = aba.get_value((i, 4))
        cnpj = str(cnpj2)

        logging.warning('Obtido CNPJ da planilha: ' + cnpj)
        print('Obtido CNPJ da planilha: ' + cnpj)
        # acessando a API publica
        url = "https://publica.cnpj.ws/cnpj/" + cnpj

        resposta = urlopen(url)
        dados = json.loads(resposta.read())

        logging.warning('Consulta no sintegra bem sucedida.')
        print('Consulta no sintegra bem sucedida.')
      
        # armazenando dados obtidos no sintegra
        ie = dados["estabelecimento"]["inscricoes_estaduais"][0]["inscricao_estadual"]
        ativo = dados["estabelecimento"]["inscricoes_estaduais"][0]["ativo"]
        tipo = dados["estabelecimento"]["tipo_logradouro"]
        endereco = dados["estabelecimento"]["logradouro"]
        numero = dados["estabelecimento"]["numero"]
        bairro = dados["estabelecimento"]["bairro"]
        complemento = dados["estabelecimento"]["complemento"]
        
        # armazenando dados no google sheets
        aba.update_value((i, 5), ie)
        aba.update_value((i, 6), ativo)
        aba.update_value((i, 7), tipo)
        aba.update_value((i, 8), endereco)
        aba.update_value((i, 9), numero)
        aba.update_value((i, 10), bairro)
        aba.update_value((i, 11), complemento)
        
        # api publica possui limite de 3 consultas por minuto
        logging.warning("Planilha atualizada com sucesso")
        print('Planilha atualizada com sucesso, aguardando 20 segundos para a próxima consulta.')

        i = i+1

        time.sleep(5)
        print(".")
        time.sleep(5)
        print(".")
        time.sleep(5)
        print(".")
        time.sleep(5)
    except:
        logging.warning("Ocorreu um erro inesperado. Tentando próximo CNPJ em 20 segundos")
        time.sleep(5)
        print(".")
        time.sleep(5)
        print(".")
        time.sleep(5)
        print(".")
        time.sleep(5)
