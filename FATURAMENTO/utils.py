import numpy as np
import pandas as pd
from datetime import datetime, timedelta, date
import calendar
import os
import math
from pprint import pprint
import requests
from time import sleep
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from  openpyxl.styles.colors import Color
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font


#dizeres com limite de 22
def get_data_NFS_root():
    columns = ['EMPRESA', 'DATA SOLICITACAO','SOLICITANTE', 'CLIENTE', 'CNPJ', 'VALOR', 'DATA FATURAMENTO',
    'CODIGO MUNICIPAL', 'DESCRICAO SERVICO','DATA INICIAL', 'DATA FINAL', 'TIPO',
    'PEDIDO', 'PROPOSTA', 'OBS', 'CONFIRMING', 'DDL', 'FIXO']
    data = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='NFS-e INPUT')
    data.columns = columns
    data = data.loc[1:, :]

    cabecalhos = data.loc[:, :'DESCRICAO SERVICO'].copy()
    cabecalhos.loc[:, 'CNPJ'] = cabecalhos.loc[:, 'CNPJ'].apply(lambda x: x.replace('.', '').replace('-', '').replace('/', '').strip() if type(x)==str else x)

    dizeres_main = data.loc[:, 'DATA INICIAL':'OBS'].copy().apply(lambda x: x.strip() if type(x)==str else x)
    for i in range(1, len(dizeres_main)+1):
        if dizeres_main.loc[i, 'TIPO'] != 'PREVENTIVA':
            try:
                for c in ['PEDIDO', 'PROPOSTA']:
                    if type(dizeres_main.loc[i, c]) == np.float64 or type(dizeres_main.loc[i, c]) == float:
                        dizeres_main.loc[i, c] = str(int(dizeres_main.loc[i, c]))
            except:
                pass
        elif type(dizeres_main.loc[i, 'PEDIDO']) == float or type(dizeres_main.loc[i, 'PEDIDO']) == np.float64:
            if math.isnan(dizeres_main.loc[i, 'PEDIDO']):
                dizeres_main.loc[i, 'PEDIDO'] = f'{datetime.now().strftime("%b/%Y")}'
    cond_pgmts = data.loc[:, 'CONFIRMING':'FIXO'].copy().apply(lambda x: x.strip() if type(x)==str else x)
    cond_pgmts = cond_pgmts.copy().apply(lambda x: int(x) if type(x)==float else x)
    for i in range(1, len(cond_pgmts)+1):
        cond_pgmt=cond_pgmts.loc[i, :].dropna()
        if cond_pgmt.empty:
            cond_pgmts.loc[i, :] = consult_cond_pgmt(cabecalhos.loc[i, 'EMPRESA'], cabecalhos.loc[i, 'CNPJ']).loc[0, :]

    return cabecalhos.reset_index().iloc[:, 1:], dizeres_main.reset_index().iloc[:, 1:], cond_pgmts.reset_index().iloc[:, 1:]

def get_data_NF_root():
    columns = ['ID', 'EMPRESA', 'DATA SOLICITACAO', 'SOLICITANTE', 'CLIENTE', 'CNPJ',
    'DATA FATURAMENTO', 'TIPO', 'OPERACAO', 'PEDIDO', 'PROPOSTA', 'OBS', 'CONFIRMING', 'DDL', 'FIXO']
    data = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='NF-e INPUT')
    data.columns = columns
    data = data.loc[1:, :]
    ####
    cabecalhos = data.loc[:, :'DATA FATURAMENTO'].copy()
    cabecalhos.loc[:, 'CNPJ'] = cabecalhos.loc[:, 'CNPJ'].apply(lambda x: x.replace('.', '').replace('-', '').replace('/', '').strip() if type(x)==str else x)
    ####
    dizeres_main = data.loc[:, 'TIPO':'OBS'].copy().apply(lambda x: x.strip() if type(x)==str else x)
    for i in range(1, len(dizeres_main)+1):
        if dizeres_main.loc[i, 'TIPO'] != 'PREVENTIVA':
            try:
                for c in ['PEDIDO', 'PROPOSTA']:
                    if type(dizeres_main.loc[i, c]) == np.float64 or type(dizeres_main.loc[i, c]) == float:
                        dizeres_main.loc[i, c] = str(int(dizeres_main.loc[i, c]))
            except:
                pass
        elif type(dizeres_main.loc[i, 'PEDIDO']) == float or type(dizeres_main.loc[i, 'PEDIDO']) == np.float64:
            if math.isnan(dizeres_main.loc[i, 'PEDIDO']):
                dizeres_main.loc[i, 'PEDIDO'] = f'{datetime.now().strftime("%b/%Y")}'
    ####
    cond_pgmts = data.loc[:, 'CONFIRMING':'FIXO'].copy().apply(lambda x: x.strip() if type(x)==str else x)
    cond_pgmts = cond_pgmts.copy().apply(lambda x: int(x) if type(x)==float else x)
    for i in range(1, len(cond_pgmts)+1):
        cond_pgmt=cond_pgmts.loc[i, :].dropna()
        if cond_pgmt.empty:
            cond_pgmts.loc[i, :] = consult_cond_pgmt(cabecalhos.loc[i, 'EMPRESA'], cabecalhos.loc[i, 'CNPJ']).loc[0, :]
    ############################################################################
    ############################################################################
    columns_materiais = ['ID', 'DESCRICAO', 'NCM', 'ST', 'QTD', 'UNIDADE', 'V_UNITARIO',
    'V_TOTAL']
    datam = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='MATERIAIS INPUT')
    datam.columns = columns_materiais
    datam = datam.loc[1:, :]
    for i in range(1, len(datam)+1):
        if type(datam.loc[i, 'NCM']) == str:
            datam.loc[i, 'NCM'] = datam.loc[i, 'NCM'].replace('.','').replace('\xa0','')
        elif type(datam.loc[i, 'NCM']) == np.float64 or type(datam.loc[i, 'NCM']) == float:
            datam.lo[i, 'NCM'] = str(int(datam.loc[i, 'NCM'])).replace('\xa0','').replace('.','')
    return cabecalhos.reset_index().iloc[:, 1:], dizeres_main.reset_index().iloc[:, 1:], cond_pgmts.reset_index().iloc[:, 1:], datam.reset_index().iloc[:, 1:]

def consult_emails(empresa, CNPJ):
    if empresa == 'CONTRUAR':
        if CNPJ.isnumeric() and len(CNPJ)>11:
            CNPJ = f"{CNPJ[:2]}.{CNPJ[2:5]}.{CNPJ[5:8]}/{CNPJ[8:12]}-{CNPJ[12:]}"
        else:
            CNPJ = f"{CNPJ[:3]}.{CNPJ[3:6]}.{CNPJ[6:9]}-{CNPJ[9:]}" #CPF
        emails = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='CONTRUAR CLIENTES')
        emails = emails.loc[emails.loc[:,'CNPJ/CPF']==CNPJ,'EMAILS']
        if len(emails)!=1:
            return None
        return emails.values[0]
    elif empresa == 'SOLAR':
        if CNPJ.isnumeric() and len(CNPJ)>11:
            CNPJ = f"{CNPJ[:2]}.{CNPJ[2:5]}.{CNPJ[5:8]}/{CNPJ[8:12]}-{CNPJ[12:]}"
        else:
            CNPJ = f"{CNPJ[:3]}.{CNPJ[3:6]}.{CNPJ[6:9]}-{CNPJ[9:]}" #CPF
        emails = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='SOLAR CLIENTES')
        emails = emails.loc[emails.loc[:,'CNPJ/CPF']==CNPJ,'EMAILS']
        if len(emails)!=1:
            return None
        return emails.values[0]
    else:
        print(f'empresa {empresa} não encontrada')
        return None

def consult_banco(empresa, CNPJ):
    if empresa == 'CONTRUAR':
        if CNPJ.isnumeric() and len(CNPJ)>11:
            CNPJ = f"{CNPJ[:2]}.{CNPJ[2:5]}.{CNPJ[5:8]}/{CNPJ[8:12]}-{CNPJ[12:]}"
        else:
            CNPJ = f"{CNPJ[:3]}.{CNPJ[3:6]}.{CNPJ[6:9]}-{CNPJ[9:]}" #CPF
        CC = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='CONTRUAR CLIENTES')
        CC = CC.loc[CC.loc[:,'CNPJ/CPF']==CNPJ,'CC']
        if len(CC)!=1:
            return None
        CC = CC.values[0]
        dados_bancarios = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='DADOS BANCARIOS')
        dados_bancarios = dados_bancarios.loc[dados_bancarios.loc[:,'CC']==CC,:]
        if len(dados_bancarios)!=1:
            return None
        return dados_bancarios.loc[:, 'BANCO':].reset_index().iloc[:, 1:]
    elif empresa == 'SOLAR':
        if CNPJ.isnumeric() and len(CNPJ)>11:
            CNPJ = f"{CNPJ[:2]}.{CNPJ[2:5]}.{CNPJ[5:8]}/{CNPJ[8:12]}-{CNPJ[12:]}"
        else:
            CNPJ = f"{CNPJ[:3]}.{CNPJ[3:6]}.{CNPJ[6:9]}-{CNPJ[9:]}" #CPF
        CC = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='SOLAR CLIENTES')
        CC = CC.loc[CC.loc[:,'CNPJ/CPF']==CNPJ,'CC']
        if len(CC)!=1:
            return None
        CC = CC.values[0]
        dados_bancarios = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='DADOS BANCARIOS')
        dados_bancarios = dados_bancarios.loc[dados_bancarios.loc[:,'CC']==CC,:]
        if len(dados_bancarios)!=1:
            return None
        return dados_bancarios.loc[:, 'BANCO':].reset_index().iloc[:, 1:]
    else:
        print(f'empresa {empresa} não encontrada')
        return None
    return

def consult_cond_pgmt(empresa, CNPJ):
    if empresa == 'CONTRUAR':
        if CNPJ.isnumeric() and len(CNPJ)>11:
            CNPJ = f"{CNPJ[:2]}.{CNPJ[2:5]}.{CNPJ[5:8]}/{CNPJ[8:12]}-{CNPJ[12:]}"
        else:
            CNPJ = f"{CNPJ[:3]}.{CNPJ[3:6]}.{CNPJ[6:9]}-{CNPJ[9:]}" #CPF
        cond_pgmt = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='CONTRUAR CLIENTES')
        cond_pgmt = cond_pgmt.loc[cond_pgmt.loc[:,'CNPJ/CPF']==CNPJ, 'CONFIRMING':'FIXO']
        if len(cond_pgmt)!=1:
            print(f'cond_pgmt referente ao CNPJ {CNPJ} não encontrada')
            return cond_pgmt
        return cond_pgmt.reset_index().iloc[:, 1:]
    elif empresa == 'SOLAR':
        if CNPJ.isnumeric() and len(CNPJ)>11:
            CNPJ = f"{CNPJ[:2]}.{CNPJ[2:5]}.{CNPJ[5:8]}/{CNPJ[8:12]}-{CNPJ[12:]}"
        else:
            CNPJ = f"{CNPJ[:3]}.{CNPJ[3:6]}.{CNPJ[6:9]}-{CNPJ[9:]}" #CPF
        cond_pgmt = pd.read_excel(r'./INPUT/fila_faturamento_root.xlsx', sheet_name='SOLAR CLIENTES')
        cond_pgmt = cond_pgmt.loc[cond_pgmt.loc[:,'CNPJ/CPF']==CNPJ, 'CONFIRMING':'FIXO']
        if len(cond_pgmt)!=1:
            print(f'cond_pgmt referente ao CNPJ {CNPJ} não encontrada')
            return cond_pgmt
        return cond_pgmt.reset_index().iloc[:, 1:]

def generate_vencimento(cond_pgmt, data_fatura):
    if type(data_fatura) == str:
        data_fatura = datetime.strptime(data_fatura, '%d/%m/%Y')
    else:
        data_fatura = data_fatura
    cond_pgmt=cond_pgmt.dropna(axis=1)
    if cond_pgmt.columns[0] == 'CONFIRMING':
        if cond_pgmt.values[0][0] == 1 or cond_pgmt.values[0][0] == '1':
            vencimento = 'CONFIRMING'
        elif cond_pgmt.columns[0] == 'CONFIRMING':
            vencimento = 'COM APRESENTAÇÃO'
    elif cond_pgmt.columns[0] == 'DDL':
        ddl = timedelta(days=int(cond_pgmt.values[0][0]))
        vencimento = (data_fatura+ddl).strftime('%d/%m/%Y')
    elif cond_pgmt.columns[0] == 'FIXO':
        if not '.next' in str(cond_pgmt.values[0][0]):
            today = int(datetime.now().strftime('%d'))
            vencimento=datetime.now()
            if today == cond_pgmt.values[0][0]:
                today+=1
            while cond_pgmt.values[0][0] != int(vencimento.strftime('%d')):
                vencimento+=timedelta(days=1)
        elif '.next' in str(cond_pgmt.values[0][0]):
            today = datetime.now()
            string_time=f'{cond_pgmt.values[0][0].split(".")[0]}/{datetime.now().strftime("%m")}/{datetime.now().strftime("%Y")}'
            vencimento = datetime.strptime(string_time, '%d/%m/%Y')+timedelta(days=1)
            while int(cond_pgmt.values[0][0].split('.')[0]) != int(vencimento.strftime('%d')):
                vencimento+=timedelta(days=1)
            vencimento = vencimento.strftime('%d/%m/%Y')
    else:
        return None
    return vencimento

def get_cliente_UF(empresa, CNPJ):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/geral/clientes/'
    payload = {}
    payload['call'] = 'ListarClientes'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
      "pagina": 1,
      "registros_por_pagina": 50,
      "apenas_importado_api": "N"
    }]
    response = requests.post(endpoint, json=payload).json()
    pgs= response['total_de_paginas']
    for pg in range(1, pgs+1):
        payload['param'] = [{
          "pagina": pg,
          "registros_por_pagina": 50,
          "apenas_importado_api": "N"
        }]
        clientes = requests.post(endpoint, json=payload).json()['clientes_cadastro']
        for cliente in clientes:
            if CNPJ == cliente['cnpj_cpf'].replace('.', '').replace('-', '').replace('/', '').strip():
                return cliente['estado']
    print('cliente não encontrado')
    return None

def consult_icms(UF_cliente):
    df = pd.read_excel('./INPUT/tabela_icms.xlsx')
    icms = df.loc[0, UF_cliente]
    return float(icms)

def get_impostos_materiais(empresa, CNPJ, UF_cliente, operacao, base_de_calculo):
    if operacao.upper() == 'PEDIDO DE VENDA':
        if empresa.upper() == 'CONTRUAR':
            impostos = {
                      "icms_sn": {
                                  'cod_sit_trib_icms_sn':"102", # 400 -> remessa
                                  'origem_icms_sn':"0",
                                  },
                      "ipi":{
                             "cod_sit_trib_ipi":"53",
                             "enquadramento_ipi": "999",
                            },
                      "pis_padrao":{
                                    "cod_sit_trib_pis":"08",
                                   },
                      "cofins_padrao":{
                                      "cod_sit_trib_cofins":"08",
                                      },
                      },
            retencoes = None
        elif empresa.upper() == 'SOLAR':
            icms = consult_icms(UF_cliente)
            impostos = {
                      "icms": {
                                  'cod_sit_trib_icms':"00",
                                  'origem_icms':"0",
                                  'aliq_icms': icms,
                                  "modalidade_icms":"0",
                                  "base_icms":base_de_calculo,
                                  "valor_icms":base_de_calculo*(icms/100)
                                  },
                      "ipi":{
                              "cod_sit_trib_ipi":"53",
                              "enquadramento_ipi": "999",
                             },
                      "pis_padrao":{
                                    "cod_sit_trib_pis":"01",
                                    "tipo_calculo_pis":"B",
                                    "aliq_pis":1.65,
                                    "base_pis":base_de_calculo,
                                    "valor_pis":base_de_calculo*(1.65/100)
                                   },
                      "cofins_padrao":{
                                      "cod_sit_trib_cofins":"01",
                                      "tipo_calculo_cofins":"B",
                                      "aliq_cofins":7.6,
                                      "base_cofins":base_de_calculo,
                                      "valor_cofins":base_de_calculo*(7.6/100)
                                      },
                      },
            values = [1.65, 7.6]
            retencoes = pd.Series(index=['PIS', 'COFINS'], data=values)
    elif operacao.upper() == 'REMESSA':
        impostos = {
                  "icms": {
                              'cod_sit_trib_icms':"41",
                              'origem_icms':"0",
                              'aliq_icms': 0,
                              "modalidade_icms":"0",
                              "base_icms":0,
                              "valor_icms":0,
                              },
                  "ipi":{
                          "cod_sit_trib_ipi":"53", #alterar para 53
                          "enquadramento_ipi": "999",
                         },
                  "pis_padrao":{
                                "cod_sit_trib_pis":"49",
                                "tipo_calculo_pis":"B",
                                "aliq_pis":0,
                                "base_pis":0,
                                "valor_pis":0
                               },
                  "cofins_padrao":{
                                  "cod_sit_trib_cofins":"49", # OUTRAS OPERAÇÕES
                                  "tipo_calculo_cofins":"B",
                                  "aliq_cofins":0,
                                  "base_cofins":0,
                                  "valor_cofins":0,
                                  },
                  },
        values = [np.nan, np.nan]
        retencoes = pd.Series(index=['PIS', 'COFINS'], data=values)
    return impostos, retencoes

def get_impostos_particularidades(CNPJ, impostos, retencoes):
    CNPJ = CNPJ.replace('.', '').replace('-', '').replace('/', '').strip()
        #CPF
    if len(CNPJ)==11:
        impostos = {
                    "cRetemINSS": "N",
                    "cRetemPIS": "N",
                    "cRetemCOFINS": "N",
                    "cRetemCSLL": "N",
                    "cRetemIRRF": "N",
                    "nAliqINSS": 0,
                    "nAliqPIS": 0,
                    "nAliqCOFINS": 0,
                    "nAliqCSLL": 0,
                    "nAliqIRRF": 0,
                    }
        values = [np.nan, np.nan, np.nan, np.nan, np.nan]
        retencoes = pd.Series(index=['INSS', 'PIS', 'COFINS', 'CSLL', 'IR'], data=values)
                #FUSSP - 11% INSS          CYMZ - EMP. VELAS
    if CNPJ == '44111698000198' or CNPJ == '34033325000192':
        impostos = {
                    "cRetemINSS": "S",
                    "cRetemPIS": "N",
                    "cRetemCOFINS": "N",
                    "cRetemCSLL": "N",
                    "cRetemIRRF": "N",
                    "nAliqINSS": 11,
                    "nAliqPIS": 0,
                    "nAliqCOFINS": 0,
                    "nAliqCSLL": 0,
                    "nAliqIRRF": 0,
                    }
        values = [11, np.nan, np.nan, np.nan, np.nan]
        retencoes = pd.Series(index=['INSS', 'PIS', 'COFINS', 'CSLL', 'IR'], data=values)
                 #ANATEL
    elif CNPJ == '02030715000201':
        impostos = {
                    "cRetemINSS": "N",
                    "cRetemPIS": "S",
                    "cRetemCOFINS": "S",
                    "cRetemCSLL": "S",
                    "cRetemIRRF": "S",
                    "nAliqINSS": 0,
                    "nAliqPIS": 0.65,
                    "nAliqCOFINS": 3,
                    "nAliqCSLL": 1,
                    "nAliqIRRF": 4.80,
                    }
        values = [np.nan, 0.65, 3, 1, 4.80]
        retencoes = pd.Series(index=['INSS', 'PIS', 'COFINS', 'CSLL', 'IR'], data=values)
                # SEADE
    elif CNPJ == '51169555000100':
        impostos = {
                    "cRetemINSS": "S",
                    "cRetemPIS": "N",
                    "cRetemCOFINS": "N",
                    "cRetemCSLL": "N",
                    "cRetemIRRF": "N",
                    "nAliqINSS": 1,
                    "nAliqPIS": 0,
                    "nAliqCOFINS": 0,
                    "nAliqCSLL": 0,
                    "nAliqIRRF": 0,
                    }
        values = [1, np.nan, np.nan, np.nan, np.nan]
        retencoes = pd.Series(index=['INSS', 'PIS', 'COFINS', 'CSLL', 'IR'], data=values)
                  #VERTE LIMÃO
    elif CNPJ == '15006151000124':
        impostos = {
                    "cRetemINSS": "S",
                    "cRetemPIS": "S",
                    "cRetemCOFINS": "S",
                    "cRetemCSLL": "S",
                    "cRetemIRRF": "N",
                    "nAliqINSS": 11,
                    "nAliqPIS": 0.65,
                    "nAliqCOFINS": 3,
                    "nAliqCSLL": 1,
                    "nAliqIRRF": 0,
                    }
        values = [11, 0.65, 3, 1, np.nan]
        retencoes = pd.Series(index=['INSS', 'PIS', 'COFINS', 'CSLL', 'IR'], data=values)
                 #FUSSP - FUNDO SOCIAL
    elif CNPJ == '44111698000198':
        impostos = {
                    "cRetemINSS": "N",
                    "cRetemPIS": "S",
                    "cRetemCOFINS": "S",
                    "cRetemCSLL": "S",
                    "cRetemIRRF": "S",
                    "nAliqINSS": 0,
                    "nAliqPIS": 0.65,
                    "nAliqCOFINS": 3,
                    "nAliqCSLL": 1,
                    "nAliqIRRF": 5,
                    }
        values = [np.nan, 0.65, 3, 1, 5]
        retencoes = pd.Series(index=['INSS', 'PIS', 'COFINS', 'CSLL', 'IR'], data=values)
    return impostos, retencoes

def get_impostos_servico(empresa, codigo_servico):
    if empresa == 'SOLAR':   #manutenção                   higienização
        if codigo_servico == '07498' or codigo_servico == '01406':
            impostos = {
                        "cRetemINSS": "N",
                        "cRetemPIS": "S",
                        "cRetemCOFINS": "S",
                        "cRetemCSLL": "S",
                        "cRetemIRRF": "N",
                        "nAliqINSS":0,
                        "nAliqPIS": 0.65,
                        "nAliqCOFINS": 3,
                        "nAliqCSLL": 1,
                        "nAliqIRRF": 0,
                        }
            values = [np.nan, 0.65, 3, 1, np.nan]
                        #      OBRA                         ELETRICA E HHIDRAULICA/OBRA
        elif codigo_servico == '07285' or codigo_servico == '01023':
            impostos = {
                        "cRetemINSS": "S",
                        "cRetemPIS": "N",
                        "cRetemCOFINS": "N",
                        "cRetemCSLL": "N",
                        "cRetemIRRF": "N",
                        "nAliqINSS": 11,
                        "nAliqPIS": 0,
                        "nAliqCOFINS": 0,
                        "nAliqCSLL": 0,
                        "nAliqIRRF": 0,
                        }
            values = [11, np.nan, np.nan, np.nan, np.nan]
                        ### RECORD RETEM PCC E INSS
    elif empresa == 'CONTRUAR':
        impostos = {
                    "cRetemINSS": "N",
                    "cRetemPIS": "N",
                    "cRetemCOFINS": "N",
                    "cRetemCSLL": "N",
                    "cRetemIRRF": "N",
                    "nAliqINSS": 0,
                    "nAliqPIS": 0,
                    "nAliqCOFINS": 0,
                    "nAliqCSLL": 0,
                    "nAliqIRRF": 0,
                    }
        values = [np.nan, np.nan, np.nan, np.nan, np.nan]
    else:
        return None, None

    retencoes = pd.Series(index=['INSS', 'PIS', 'COFINS', 'CSLL', 'IR'], data=values)
    return impostos, retencoes

def get_API(empresa):
    if empresa.upper() == 'SOLAR':
        OMIE_APP_KEY = '1842612490719'
        OMIE_APP_SECRET = '9703e43b493ed4e221c8c991abe4c087'
    elif empresa.upper() == 'CONTRUAR':
        OMIE_APP_KEY = '1842673823991'
        OMIE_APP_SECRET = 'fbb03ec8c37377049b77b0b63390bdfc'
    else:
        return None, None
    return OMIE_APP_KEY, OMIE_APP_SECRET

def consultar_os(pedido, proposta, nCodCli, OMIE_APP_KEY, OMIE_APP_SECRET):
    endpoint = 'https://app.omie.com.br/api/v1/servicos/os/'
    payload = {}
    payload['call'] = 'ListarOS'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{'pagina':1, 'registros_por_pagina':50, "apenas_importado_api":"N"}]
    total_de_paginas = requests.post(endpoint, json=payload).json()['total_de_paginas']
    pendentes = []
    faturados = []
    cancelados = []
    if nCodCli == None:
        for pg in range(1, total_de_paginas+1):
            print(0)
            payload['param'] = [{'pagina':pg, 'registros_por_pagina':50, "apenas_importado_api":"N"}]
            response = requests.post(endpoint, json=payload).json()
            cadastros = response['osCadastro']
            l = []
            for cadastro in cadastros:
                info_ad = cadastro['InformacoesAdicionais']
                info_cad = cadastro['InfoCadastro']
                l.append([cadastro])
                v=0
                if 'cNumPedido' in info_ad:
                    if info_ad['cNumPedido'] == pedido:
                        if info_cad['cCancelada']!='S' and info_cad['cFaturada']!='S':
                            print('1.PENDENTE', pedido)
                            pendentes.append(cadastro)
                            v=1
                        elif info_cad['cCancelada']!='S' and info_cad['cFaturada']=='S':
                            print('1.FATURADO', pedido)
                            faturados.append(cadastro)
                            v=1
                        elif info_cad['cCancelada']=='S':
                            print('1.CANCELADO', pedido)
                            cancelados.append(cadastro)
                            v=1
                if 'cNumContrato' in info_ad and v!=1:
                    if info_ad['cNumContrato'] == proposta:
                        if info_cad['cCancelada']!='S' and info_cad['cFaturada']!='S':
                            print('2.PENDENTE', proposta)
                            pendentes.append(cadastro)
                        elif info_cad['cCancelada']!='S' and info_cad['cFaturada']=='S':
                            print('2.FATURADO', proposta)
                            faturados.append(cadastro)
                        elif info_cad['cCancelada']=='S':
                            print('2.CANCELADO', proposta)
                            cancelados.append(cadastro)
        cadastros_os = {'pendentes':pendentes, 'faturados':faturados, 'cancelados':cancelados}
        if cadastros_os == {'pendentes':[], 'faturados':[], 'cancelados':[]}:
            return None
        else:
            return cadastros_os
    elif nCodCli != None:
        for pg in range(1, total_de_paginas+1):
            print(0)
            payload['param'] = [{'pagina':pg, 'registros_por_pagina':50, "apenas_importado_api":"N"}]
            response = requests.post(endpoint, json=payload).json()
            cadastros = response['osCadastro']
            l = []
            for cadastro in cadastros:
                info_ad = cadastro['InformacoesAdicionais']
                info_cad = cadastro['InfoCadastro']
                l.append([cadastro])
                if cadastro['Cabecalho']['nCodCli'] == nCodCli:
                    if 'cNumPedido' in info_ad:
                        if info_ad['cNumPedido'] == pedido:
                            if info_cad['cCancelada']!='S' and info_cad['cFaturada']!='S':
                                print('1.PENDENTE', pedido)
                                pendentes.append(cadastro)
                            elif info_cad['cCancelada']!='S' and info_cad['cFaturada']=='S':
                                print('1.FATURADO', pedido)
                                faturados.append(cadastro)
                            elif info_cad['cCancelada']=='S':
                                print('1.CANCELADO', pedido)
                                cancelados.append(cadastro)
        cadastros_os = {'pendentes':pendentes, 'faturados':faturados, 'cancelados':cancelados}
        if cadastros_os == {'pendentes':[], 'faturados':[], 'cancelados':[]}:
            return None
        else:
            return cadastros_os

def consultar_pv(pedido, proposta, nCodCli, OMIE_APP_KEY, OMIE_APP_SECRET):
    endpoint = 'https://app.omie.com.br/api/v1/produtos/pedido/'
    payload = {}
    payload['call'] = 'ListarPedidos'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{'pagina':1, 'registros_por_pagina':50, "apenas_importado_api":"N"}]
    total_de_paginas = requests.post(endpoint, json=payload).json()['total_de_paginas']
    pendentes = []
    faturados = []
    cancelados = []
    if nCodCli == None:
        for pg in range(1, total_de_paginas+1):
            print(0)
            payload['param'] = [{'pagina':pg, 'registros_por_pagina':50, "apenas_importado_api":"N"}]
            response = requests.post(endpoint, json=payload).json()
            cadastros = response['pedido_venda_produto']
            l = []
            for cadastro in cadastros:
                info_ad = cadastro['informacoes_adicionais']
                info_cad = cadastro['infoCadastro']
                l.append([cadastro])
                v=0
                if 'numero_pedido_cliente' in info_ad:
                    if info_ad['numero_pedido_cliente'] == pedido:
                        if info_cad['cancelado']!='S' and info_cad['faturado']!='S':
                            print('1.PENDENTE', pedido)
                            pendentes.append(cadastro)
                            v=1
                        elif info_cad['cancelado']!='S' and info_cad['faturado']=='S':
                            print('1.FATURADO', pedido)
                            faturados.append(cadastro)
                            v=1
                        elif info_cad['cancelado']=='S':
                            print('1.CANCELADO', pedido)
                            cancelados.append(cadastro)
                            v=1
                if 'numero_contrato' in info_ad and v!=1:
                    if info_ad['numero_contrato'] == proposta:
                        if info_cad['cancelado']!='S' and info_cad['faturado']!='S':
                            print('2.PENDENTE', proposta)
                            pendentes.append(cadastro)
                        elif info_cad['cancelado']!='S' and info_cad['faturado']=='S':
                            print('2.FATURADO', proposta)
                            faturados.append(cadastro)
                        elif info_cad['cancelado']=='S':
                            print('2.CANCELADO', proposta)
                            cancelados.append(cadastro)
        cadastros_pv = {'pendentes':pendentes, 'faturados':faturados, 'cancelados':cancelados}
        if cadastros_pv == {'pendentes':[], 'faturados':[], 'cancelados':[]}:
            return None
        else:
            return cadastros_pv
    elif nCodCli != None:
        for pg in range(1, total_de_paginas+1):
            print(0)
            payload['param'] = [{'pagina':pg, 'registros_por_pagina':50, "apenas_importado_api":"N"}]
            response = requests.post(endpoint, json=payload).json()
            cadastros = response['pedido_venda_produto']
            for cadastro in cadastros:
                info_ad = cadastro['informacoes_adicionais']
                info_cad = cadastro['infoCadastro']
                if cadastro['cabecalho']['codigo_cliente'] == nCodCli:
                    if 'numero_pedido_cliente' in info_ad:
                        if info_ad['numero_pedido_cliente'] == pedido and pedido!='':
                            if info_cad['cancelado']!='S' and info_cad['faturado']!='S':
                                print('1.PENDENTE', pedido)
                                pendentes.append(cadastro)
                            elif info_cad['cancelado']!='S' and info_cad['faturado']=='S':
                                print('1.FATURADO', pedido)
                                faturados.append(cadastro)
                            elif info_cad['cancelado']=='S':
                                print('1.CANCELADO', pedido)
                                cancelados.append(cadastro)
        cadastros_pv = {'pendentes':pendentes, 'faturados':faturados, 'cancelados':cancelados}
        if cadastros_pv == {'pendentes':[], 'faturados':[], 'cancelados':[]}:
            return None
        else:
            return cadastros_pv

def listar_pv_faturar(empresa):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/produtos/pedido/'
    payload = {}
    payload['call'] = 'ListarPedidos'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{'pagina':1, 'registros_por_pagina':50, "apenas_importado_api":"N"}]
    response = requests.post(endpoint, json=payload).json()
    pgs = response['total_de_paginas']
    lst_pv_faturar = []
    for pg in range(1, pgs+1):
        payload['param'] = [{'pagina':pg, 'registros_por_pagina':50, "apenas_importado_api":"N"}]
        response = requests.post(endpoint, json=payload).json()
        for pv in response['pedido_venda_produto']:
            if pv['cabecalho']['etapa'] == '50':
                lst_pv_faturar.append(pv)
    return lst_pv_faturar

def get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET):
    endpoint = 'https://app.omie.com.br/api/v1/geral/clientes/'
    payload = {}
    payload['call'] = 'ListarClientes'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
      "pagina": 1,
      "registros_por_pagina": 50,
      "apenas_importado_api": "N"
    }]
    response = requests.post(endpoint, json=payload).json()
    pgs= response['total_de_paginas']
    for pg in range(1, pgs+1):
        payload['param'] = [{
          "pagina": pg,
          "registros_por_pagina": 50,
          "apenas_importado_api": "N"
        }]
        response = requests.post(endpoint, json=payload).json()['clientes_cadastro']
        for cliente in response:
            if CNPJ == cliente['cnpj_cpf'].replace('.', '').replace('-', '').replace('/', '').strip():
                return cliente['codigo_cliente_omie']
    return None

def get_nCodCC(Num_CC, OMIE_APP_KEY, OMIE_APP_SECRET):
    endpoint = 'https://app.omie.com.br/api/v1/geral/contacorrente/'
    payload = {}
    payload['call'] = 'ListarContasCorrentes'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
      "pagina": 1,
      "registros_por_pagina": 100,
      "apenas_importado_api": "N"
    }]
    response = requests.post(endpoint, json=payload).json()['ListarContasCorrentes']
    for CC in response:
        if 'numero_conta_corrente' in CC.keys():
            if Num_CC in CC['numero_conta_corrente']:
                return CC['nCodCC']
    print('\n\n\nCONTA CORRENTE NÃO ENCONTRADA!\n\n\n')
    return None

def generate_dizeres(cabecalho, dizeres_main, cond_pgmt, vencimento, dados_bancarios, retencoes):
    """
USAGE:
x=0
dizeres = generate_dizeres(cabecalhos.loc[x, :], dizeres_main.loc[x, :], cond_pgmnt.loc[x, :], dados_bancarios.loc[x, :], retencoes.loc[x, :])

"""
    dizeres_main = dizeres_main.fillna('NULL')
    date0 = dizeres_main['DATA INICIAL']
    date1 = dizeres_main['DATA FINAL']
    if date0 != 'NULL' and date0 == date1:
        if str(date0).isalpha():
            executado = f"""
    EXECUTADO EM {dizeres_main['DATA INICIAL']}"""
        else:
            executado = f"""
    EXECUTADO: {date0.strftime('%d/%m/%Y')}"""
    elif date0 != 'NULL' and date0 != date1 and type(date0)!=str:
        executado = f"""
    EXECUÇÃO: {date0.strftime('%d/%m/%Y')} A {date1.strftime('%d/%m/%Y')}"""
    elif date0 != 'NULL' and date0 != date1:
        executado = f"""
    EXECUÇÃO: {date0}/{date1}"""
    else:
        executado=''
    ############################################################################
    obs = dizeres_main['OBS']
    if obs != 'NULL':
        obss = obs.split('|')
        dizeres_obs = ""
        for ob in obss:
            dizeres_obs += f"""
{ob}"""
    else:
        dizeres_obs = ''
    ############################################################################
    cond_pgmt=cond_pgmt.fillna('NULL')
    CONFIRMING = cond_pgmt['CONFIRMING'].values[0]
    DDL = cond_pgmt['DDL'].values[0]
    FIXO = cond_pgmt['FIXO'].values[0]
    VENCIMENTO = vencimento
    if DDL != 'NULL':
        today = datetime.now()
        venc = (today+timedelta(days=int(cond_pgmt['DDL'])) ).strftime('%d/%m/%Y')
        diz_vencimento = f"""
VENCIMENTO: {venc} ({DDL} DDL)"""
    elif FIXO != 'NULL':
        diz_vencimento = f"""
VENCIMENTO: {VENCIMENTO}"""
    elif CONFIRMING != 'NULL':
        diz_vencimento = ''
    elif VENCIMENTO != None:
        diz_vencimento = f"""
VENCIMENTO: {VENCIMENT}"""
    else:
        diz_vencimento = ''
    ############################################################################
    dizeres_main=dizeres_main.fillna('NULL')
    TIPO = dizeres_main['TIPO']
    PEDIDO = dizeres_main['PEDIDO']
    if type(PEDIDO) == np.float64 or type(PEDIDO) == np.float32:
        PEDIDO = int(PEDIDO)


    PROPOSTA = dizeres_main['PROPOSTA']
    if TIPO == 'CORRETIVA' or TIPO == 'OBRA':
        if PEDIDO != 'NULL':
            pedido = f"""
PEDIDO: {PEDIDO}"""
        else:
            pedido = ''
        if PROPOSTA != 'NULL':
            proposta = f"""
PROPOSTA: {PROPOSTA}"""
        else:
            proposta = ''
        codigos = dizeres_obs+pedido+proposta
    elif TIPO == 'PREVENTIVA':
        if PEDIDO != 'NULL':
            pedido = f"""
PEDIDO: {PEDIDO}"""
        else:
            pedido = ''
        if PROPOSTA != 'NULL':
            contrato = f"""
CONFORME CONTRATO: {PROPOSTA}"""
        else:
            contrato = """
CONFORME CONTRATO"""
        codigos = dizeres_obs+pedido+contrato
    else:
        codigos=""
    ############################################################################
    dados_bancarios=dados_bancarios.fillna('NULL')
    banco = dados_bancarios['BANCO'].values[0]
    ag = dados_bancarios['AGENCIA'].values[0]
    CC = dados_bancarios['CC'].values[0]
    PIX = dados_bancarios['PIX'].values[0]
    if cond_pgmt['CONFIRMING'].values[0] == 'NULL':
        if banco != 'NULL' and ag != 'NULL' and CC != 'NULL':
            dizeres_pagamento = f"""
FORMA DE PAGAMENTO: DEPÓSITO BANCÁRIO
DADOS BANCÁRIOS:
BANCO {banco}, AG. {ag}, C/C. {CC}"""
        if PIX != 'NULL':
            pix = f"""
PIX: {PIX}
"""
        else:
            pix=''
        dizeres_pagamento+=pix
    else:
        if banco != 'NULL' and ag != 'NULL' and CC != 'NULL':
            dizeres_pagamento = f"""
FORMA DE PAGAMENTO: DEPÓSITO BANCÁRIO
DADOS BANCÁRIOS:
BANCO {banco}, AG. {ag}, C/C. {CC}"""
        if PIX != 'NULL':
            pix = f"""
PIX: {PIX}
"""
        else:
            pix=''
        dizeres_pagamento+=pix
    ############################################################################
    empresa = cabecalho.EMPRESA
    if empresa == 'SOLAR':
        diz_retencao = f"""
RETENÇÕES:"""
        for retencao, percent in retencoes.items():
            if not math.isnan(percent):
                diz_retencao+=f"""
{retencao}: {str(percent).replace('.', ',')}% """
        if len(diz_retencao) == 11:
            diz_retencao=''
    elif empresa == 'CONTRUAR':
        diz_retencao="""
OPTANTE PELO SIMPLES NACIONAL NÃO POSSUI RETENÇÕES"""
    ############################################################################
    dizeres=executado+diz_vencimento+codigos+dizeres_pagamento+diz_retencao
    return dizeres

def generate_dizeres_PV(cabecalho, dizeres_main, cond_pgmt, vencimento, dados_bancarios, retencoes):
    ############################################################################
    dizeres_main = dizeres_main.fillna('NULL')
    obs = dizeres_main['OBS']
    if obs != 'NULL':
        obss = obs.split('|')
        dizeres_obs = ""
        for ob in obss:
            dizeres_obs += f"""
{ob}"""
    else:
        dizeres_obs = ''

    ############################################################################
    TIPO = dizeres_main['TIPO']
    PEDIDO = dizeres_main['PEDIDO']
    if type(PEDIDO) == np.float64 or type(PEDIDO) == np.float32:
        PEDIDO = int(PEDIDO)

    PROPOSTA = dizeres_main['PROPOSTA']
    if TIPO == 'CORRETIVA' or TIPO == 'OBRA':
        if PEDIDO != 'NULL':
            pedido = f"""
PEDIDO: {PEDIDO}"""
        else:
            pedido = ''
        if PROPOSTA != 'NULL':
            proposta = f"""
PROPOSTA: {PROPOSTA}"""
        else:
            proposta = ''
        codigos = dizeres_obs+pedido+proposta
    elif TIPO == 'PREVENTIVA':
        if PEDIDO != 'NULL':
            pedido = f"""
PEDIDO: {PEDIDO}"""
        else:
            pedido = ''
        if PROPOSTA != 'NULL':
            contrato = f"""
CONFORME CONTRATO: {PROPOSTA}"""
        else:
            contrato = """
CONFORME CONTRATO"""
        codigos = dizeres_obs+pedido+contrato
    else:
        codigos=""
    ############################################################################
    dados_bancarios=dados_bancarios.fillna('NULL')
    banco = dados_bancarios['BANCO'].values[0]
    ag = dados_bancarios['AGENCIA'].values[0]
    CC = dados_bancarios['CC'].values[0]
    PIX = dados_bancarios['PIX'].values[0]
    if math.isnan(cond_pgmt['CONFIRMING'].values[0]):
        if banco != 'NULL' and ag != 'NULL' and CC != 'NULL':
            dizeres_pagamento = f"""
FORMA DE PAGAMENTO: DEPÓSITO BANCÁRIO
DADOS BANCÁRIOS:
BANCO {banco}, AG. {ag}, C/C. {CC}"""
        if PIX != 'NULL':
            pix = f"""
PIX: {PIX}
"""
        else:
            pix=''
        dizeres_pagamento+=pix
    else:
        if banco != 'NULL' and ag != 'NULL' and CC != 'NULL':
            dizeres_pagamento = f"""
FORMA DE PAGAMENTO: CONFIRMING
DADOS BANCÁRIOS:
BANCO {banco}, AG. {ag}, C/C. {CC}"""
        if PIX != 'NULL':
            pix = f"""
PIX: {PIX}
"""
        else:
            pix=''
        dizeres_pagamento+=pix
    ############################################################################
    empresa = cabecalho.EMPRESA
    if empresa == 'SOLAR':
        diz_retencao = f"""
RETENÇÕES:"""
        for retencao, percent in retencoes.items():
            if not math.isnan(percent):
                diz_retencao+=f"""
{retencao}: {str(percent).replace('.', ',')}% """
        if len(diz_retencao) == 11:
            diz_retencao=''
    elif empresa == 'CONTRUAR':
        diz_retencao="""
OPTANTE PELO SIMPLES NACIONAL NÃO POSSUI RETENÇÕES"""
    ############################################################################
    dizeres=codigos+dizeres_pagamento+diz_retencao
    return dizeres

def get_Status_OS(nCodOS, empresa):
    endpoint = 'https://app.omie.com.br/api/v1/servicos/os/'
    payload = {}
    payload['call'] = 'StatusOS'
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{"nCodOS": nCodOS,}]
    response = requests.post(endpoint, json=payload).json()
    return response

def get_Nat(operacao):
    if 'REMESSA' in operacao:
        natopr = 'NOTA FISCAL DE REMESSA'
    elif 'VENDA' in operacao:
        natopr = 'NOTA FISCAL DE VENDA'
    elif 'DEVOL' in operacao:
        natopr = 'NOTA FISCAL DE DEVOLUÇÃO'
    return natopr

def get_nfe(empresa, codigo_pedido):
    endpoint = 'https://app.omie.com.br/api/v1/produtos/nfconsultar/'
    payload = {}
    payload['call'] = 'ListarNF'
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{"pagina": 1, "registros_por_pagina": 50, "apenas_importado_api": "N"}]
    response = requests.post(endpoint, json=payload).json()
    print(response)
    pgs = response['total_de_paginas']
    lista = []
    for pg in range(1, pgs+1):
        payload['param'] = [{"pagina": pg, "registros_por_pagina": 50,  "apenas_importado_api": "N",}]
        response = requests.post(endpoint, json=payload).json()
        nfes = response['nfCadastro']
        for nfe in nfes:
            if nfe['compl']['nIdPedido'] == codigo_pedido:
                return nfe
    return None

def get_nfse(empresa, nCodOS):
    endpoint = 'https://app.omie.com.br/api/v1/servicos/nfse/'
    payload = {}
    payload['call'] = 'ListarNFSEs'
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{"nPagina": 1, "nRegPorPagina": 50}]
    response = requests.post(endpoint, json=payload).json()
    pgs = response['nTotPaginas']
    lista = []
    for pg in range(1, pgs+1):
        payload['param'] = [{"nPagina": pg, "nRegPorPagina": 50}]
        response = requests.post(endpoint, json=payload).json()
        nfses = response['nfseEncontradas']
        for nfse in nfses:
            if nfse['OrdemServico']['nCodigoOS'] == nCodOS:
                return nfse
    return None

def get_cliente(empresa, CNPJ):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/geral/clientes/'
    payload = {}
    payload['call'] = 'ListarClientes'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{'pagina':1, 'registros_por_pagina':100, 'apenas_importado_api':'N'},]
    response = requests.post(endpoint, json=payload).json()
    pgs = response['total_de_paginas']
    for pg in range(1, pgs+1):
        payload['param'] = [{'pagina':pg, 'registros_por_pagina':100, 'apenas_importado_api':'N'},]
        response = requests.post(endpoint, json=payload).json()
        clientes = response['clientes_cadastro']
        for cliente in clientes:
            if cliente['cnpj_cpf'].replace('.', '').replace('-', '').replace('/', '').strip() == CNPJ.replace('.', '').replace('-', '').replace('/', '').strip():
                return cliente
    return None

def get_material(empresa, material):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/geral/produtos/'
    payload = {}
    payload['call'] = 'ListarProdutos'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                          "pagina": 1,
                          "registros_por_pagina": 45,
                          "apenas_importado_api": "N",
                          "filtrar_apenas_omiepdv": "N"
                        },]
    response = requests.post(endpoint, json=payload).json()
    pgs = response['total_de_paginas']
    for pg in range(1, pgs+1):
        print(0.)
        payload['param'] = [{
                              "pagina": pg,
                              "registros_por_pagina": 45,
                              "apenas_importado_api": "N",
                              "filtrar_apenas_omiepdv": "N"
                            },]
        response = requests.post(endpoint, json=payload).json()
        materiais_cadastrados = response['produto_servico_cadastro']
        for m_cadastrado in materiais_cadastrados:
            if m_cadastrado['descricao'] == material['DESCRICAO']:
                return m_cadastrado
    return None

def get_cfop(operacao, UF_cliente, ST):
    if UF_cliente=='SP':
        cfop0='5.'
    else:
        cfop0='6.'
    if 'VENDA' in operacao:
        if int(ST) == 1:
            cfop1 = '405'# cpf = "108"
        elif int(ST) == 0:
            cfop1='102'# cpf = "108"
    elif operacao == 'REMESSA':
        if int(ST) == 1:
            cfop1 = '415'
        elif int(ST) == 0:
            cfop1 = '949'
    elif 'DEVOL' in operacao:
        if int(ST) == 1:
            cfop = '411'
        elif int(ST) == 0:
            cfop = '202'
    else:
        print(f'operacao de nota {operacao} não identificado')
    cfop = cfop0+cfop1
    return cfop

def consultar_material(empresa, codigo):

    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/geral/produtos/'
    payload = {}
    payload['call'] = 'ConsultarProduto'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                          "codigo_produto":0,
                          "codigo_produto_integracao": "",
                          "codigo": codigo,
                        }]
    response = requests.post(endpoint, json=payload).json()
    return response

def listar_categorias(empresa):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/geral/categorias/'
    payload = {}
    payload['call'] = 'ListarCategorias'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                          "pagina": 1,
                          "registros_por_pagina": 50
                        }]
    response = requests.post(endpoint, json=payload).json()
    pgs = response['total_de_paginas']
    lst = []
    for pg in range(1, pgs+1):
        payload['param'] = [{
                              "pagina": pg,
                              "registros_por_pagina": 50
                            }]
        response = requests.post(endpoint, json=payload).json()
        for categoria in response['categoria_cadastro']:
            if categoria['descricao'] != '&lt;Disponível&gt;' and categoria['conta_despesa']=='S':
                pprint(f"CODIGO: {categoria['codigo']}, DESC: {categoria['descricao']}")
                lst.append({'CODIGO':categoria['codigo'], 'DESC':categoria['descricao']})
    return lst

def listar_tipos_doc(empresa):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/geral/tiposdoc/'
    payload = {}
    payload['call'] = 'PesquisarTipoDocumento'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                          "codigo": ""
                        }]
    response = requests.post(endpoint, json=payload).json()
    tipos = response['tipo_documento_cadastro']
    lst=[]
    for tipo in tipos:
        lst.append({'CODIGO':tipo['codigo'] ,'DESCRICAO':tipo['descricao']})
    return lst

def find_client_by(nCodCli, empresa):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/geral/clientes/'
    payload = {}
    payload['call'] = 'ConsultarCliente'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                          "codigo_cliente_omie": nCodCli,
                          "codigo_cliente_integracao": ""
                        },]
    response = requests.post(endpoint, json=payload).json()
    return response

def get_contas_pagar(empresa):
    global df
    df = pd.DataFrame()
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/financas/contapagar/'
    payload = {}
    payload['call'] = 'ListarContasPagar'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                        "pagina": 1,
                        "registros_por_pagina": 50,
                        "apenas_importado_api": "N"
                        },]
    response = requests.post(endpoint, json=payload).json()
    pgs = response['total_de_registros']
    for pg in range(1, pgs+1):
        payload['param'] = [{
                            "pagina": pg,
                            "registros_por_pagina": 50,
                            "apenas_importado_api": "N"
                            },]
        response = requests.post(endpoint, json=payload).json()
        if 'conta_pagar_cadastro' in response:
            cadastros = response['conta_pagar_cadastro']
            for cadastro in cadastros:
                cli = find_client_by(cadastro['codigo_cliente_fornecedor'], empresa)
                if 'numero_documento_fiscal' in cadastro:
                    num_fiscal = cadastro['numero_documento_fiscal']
                else:
                    num_fiscal = np.nan
                if 'razao_social' in cli:
                    cli_razao = cli['razao_social']
                else:
                    cli_razao = cadastro['codigo_cliente_fornecedor']
                dct={
                    'cliente': cli_razao,
                    'nCodCli': cadastro['codigo_cliente_fornecedor'],
                    'numero_fiscal': num_fiscal,
                    'emissao': cadastro['data_emissao'],
                    'vencimento': cadastro['data_vencimento'],
                    'num_parcela': cadastro['numero_parcela'],
                    'valor_documento': cadastro['valor_documento'],
                    'codigo_lancamento_omie': cadastro['codigo_lancamento_omie'],
                    'status': cadastro['status_titulo'],
                    'doc': cadastro['codigo_tipo_documento'],
                    }
                pprint(f"{cli_razao}: {num_fiscal}")
                df = df.append(dct, ignore_index=True)
    return df
################################################################################
################################################################################
################################################################################
# imprimir docs fiscais a partir do xml https://github.com/erpbrasil/erpbrasil.edoc.pdf

def apurar_NFe_saida(empresa, mes, d0, d1):
    endpoint_contador = 'https://app.omie.com.br/api/v1/contador/xml/'
    payload = {}
    payload['call'] = 'ListarDocumentos'
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{'nPagina':1, 'nRegPorPagina':50, 'cModelo':55, 'cOperacao':'1', 'dEmiInicial':d0, 'dEmiFinal':d1}]
    response = requests.post(endpoint_contador, json=payload).json()
    nTotPaginas = response['nTotPaginas']
    xmls = {}
    id = 1
    for pg in range(1, nTotPaginas+1):                      #55=NF-e | 99=NFS-e  #0=entrada - 1=saida
        payload['param'] = [{'nPagina':pg, 'nRegPorPagina':50, 'cModelo':55, 'cOperacao':'1', 'dEmiInicial':d0, 'dEmiFinal':d1}]
        response = requests.post(endpoint_contador, json=payload).json()
        notas = response['documentosEncontrados']
        for nota in notas:
            print(nota['cStatus'])
            if nota['cStatus'] == '00':
                status = 'APROVADA'
            elif nota['cStatus'] == '10':
                status = 'CANCELADA'
            elif nota['cStatus'] == '20':
                status = 'DENEGADA'
            elif nota['cStatus'] == '30': # <-- NÃO RETORNA MENHUM XML COM ESSE STATUS
                status = 'DEVOLUÇÃO'
            elif nota['cStatus'] == '40':
                status = 'INUTILIZADA'
            string_xml = nota['cXml'].replace('&lt;', '<').replace('&quot;', '"').replace('&gt;', '>')
            xml = BeautifulSoup(string_xml, 'xml')
            with open(f"./APURACAO {mes}/{empresa}/SAIDA/XML/{nota['dEmissao'].replace('/','_')} {nota['nNumero']} {status}.xml", 'w') as f:
                f.write(string_xml)
    return

def apurar_NFe_entrada(empresa, mes, d0, d1):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint_contador = 'https://app.omie.com.br/api/v1/contador/xml/'
    payload = {}
    payload['call'] = 'ListarDocumentos'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{'nPagina':1, 'nRegPorPagina':20, 'cModelo':'55', 'cOperacao':'0', 'dEmiInicial':d0, 'dEmiFinal':d1}]
    response = requests.post(endpoint_contador, json=payload).json()
    pgs = response['nTotPaginas']
    xmls = {}
    id = 1
    for pg in range(1, pgs+1):
        payload['param'] = [{'nPagina':pg, 'nRegPorPagina':20, 'cModelo':'55', 'cOperacao':'0', 'dEmiInicial':d0, 'dEmiFinal':d1}]
        r = requests.post(endpoint_contador, json=payload)
        notas = r.json()['documentosEncontrados']
        for i in range(0, len(notas)):
            string_xml = notas[i]['cXml'].replace('&lt;', '<').replace('&quot;', '"').replace('&gt;', '>')
            xml = BeautifulSoup(string_xml, 'xml')
            with open(f"./APURACAO {mes}/{empresa.upper()}/ENTRADA/XML/{xml.dhRecbto.string[:10]}_{str(notas[i]['nNumero'])}.xml", 'w') as f:
                f.write(string_xml)
            id+=1
    return

def apuracao_NFe_MODEL(empresa):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint_contador = 'https://app.omie.com.br/api/v1/contador/xml/'
    payload = {}
    payload['call'] = 'ListarDocumentos'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{'nPagina':1, 'nRegPorPagina':50, 'cModelo':'99', 'cOperacao':'1', 'dEmiInicial':'01/01/2022', 'dEmiFinal':'27/02/2022'}]
    response = requests.post(endpoint_contador, json=payload).json()
    pgs = response['nTotPaginas']
    xmls = {}
    id = 1
    for pg in range(1, pgs+1):
        payload['param'] = [{'nPagina':pg, 'nRegPorPagina':50, 'cModelo':'99', 'cOperacao':'1', 'dEmiInicial':'01/01/2022', 'dEmiFinal':'27/02/2022'}]
        response = requests.post(endpoint_contador, json=payload).json()
        notas = response['documentosEncontrados']
        for i in range(0, len(notas)):
            string_xml = notas[i]['cXml'].replace('&lt;', '<').replace('&quot;', '"').replace('&gt;', '>')
            xml = BeautifulSoup(string_xml, 'xml')
            tomador = xml.find_all('CPFCNPJTomador')
            if len(tomador)>=1:
                NF_CNPJ_tomador = tomador[0].find_all('CNPJ')
                if len(NF_CNPJ_tomador)>=1:
                    if '11137051' in NF_CNPJ_tomador[0].string:
                        with open(f"./xml_boticario/{xml.DataEmissaoRPS.string}_{str(notas[i]['nNumero'])}.xml", 'w') as f:
                            f.write(string_xml)
                        id+=1
    return

    #numero do mes

def apurar_NFe(mes_n, ano=None):
    meses = ['JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
    mes = meses[mes_n-1]
    empresas = ['SOLAR', 'CONTRUAR']
    ################
    d = date.today()
    if ano == None:
        ano = d.year
    else:
        ano = ano
    d0 = d.replace(day=1, month=mes_n, year=ano).strftime('%d/%m/%Y')
    d1 = d.replace(day=calendar.monthrange(ano, mes_n)[1], month=mes_n, year=ano).strftime('%d/%m/%Y')
    ############################################################################
    os.mkdir(f'./APURACAO {mes}/')
    for empresa in empresas:
        os.mkdir(f'./APURACAO {mes}/{empresa}/')
        os.mkdir(f'./APURACAO {mes}/{empresa}/ENTRADA/')
        os.mkdir(f'./APURACAO {mes}/{empresa}/ENTRADA/DANFE/')
        os.mkdir(f'./APURACAO {mes}/{empresa}/ENTRADA/XML/')
        apurar_NFe_entrada(empresa, mes, d0, d1)
        os.mkdir(f'./APURACAO {mes}/{empresa}/SAIDA/')
        os.mkdir(f'./APURACAO {mes}/{empresa}/SAIDA/DANFE/')
        os.mkdir(f'./APURACAO {mes}/{empresa}/SAIDA/XML/')
        apurar_NFe_saida(empresa, mes, d0, d1)
    return
