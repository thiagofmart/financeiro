from utils import *


def cadastrar_material(empresa, material):
    ID_df = pd.read_excel(f'./INPUT/ID/IDmaterial{empresa.upper()}.xlsx', sheet_name=empresa.upper())
    ID = int(ID_df.loc[len(ID_df)-1, 'ID']+1)#@
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    NCM = f"{str(material['NCM'])[:4]}.{str(material['NCM'])[4:6]}.{str(material['NCM'])[6:]}"
    ############################################################################
    endpoint = 'https://app.omie.com.br/api/v1/geral/produtos/'
    payload = {}
    payload['call'] = 'UpsertProduto'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                          "codigo_produto_integracao": f"SOLR {ID}",
                          "codigo": f'SOLR {ID}',
                          "descricao": material['DESCRICAO'],
                          "unidade": material['UNIDADE'],
                          "ncm": NCM,
                        },]
    response = requests.post(endpoint, json=payload)
    ID_df = ID_df.append({'ID':ID}, ignore_index=True)
    ID_df.to_excel(f'./INPUT/ID/IDmaterial{empresa.upper()}.xlsx', sheet_name=empresa.upper(), index=False)
    return response

def criar_pv(cabecalho, dizeres_main, cond_pgmt, vencimento, materiais, valor):
    """
TEST

empresa = cabecalho.EMPRESA#@
CNPJ = cabecalho.CNPJ#@
OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)#@
codigo_cliente = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)#@
cCodIntPV_df = pd.read_excel(f'./INPUT/ID/IDpedido_de_venda{empresa.upper()}.xlsx', sheet_name=empresa.upper())
cCodIntPV = cCodIntPV_df.loc[len(cCodIntPV_df)-1, 'ID']+1#@
dados_bancarios = consult_banco(empresa, CNPJ)#@
Num_CC = dados_bancarios['CC'].values[0]#@
nCodCC = get_nCodCC(Num_CC, OMIE_APP_KEY, OMIE_APP_SECRET)#@
# CONTINUAR PELOS DIZERES
dizeres = generate_dizeres_PV(cabecalho, dizeres_main, cond_pgmt, vencimento, dados_bancarios, retencoes)
dizeres = ""
emails = consult_emails(empresa, CNPJ)#@
    """
    empresa = cabecalho.EMPRESA#@
    CNPJ = cabecalho.CNPJ#@
    operacao = dizeres_main.OPERACAO
    natopr = get_Nat(operacao)
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)#@
    codigo_cliente = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)#@
    cCodIntPV_df = pd.read_excel(f'./INPUT/ID/IDpedido_de_venda{empresa.upper()}.xlsx', sheet_name=empresa.upper())
    cCodIntPV = cCodIntPV_df.loc[len(cCodIntPV_df)-1, 'ID']+1#@
    if 'VENDA' in operacao:
        dados_bancarios = consult_banco(empresa, CNPJ)#@
        Num_CC = dados_bancarios['CC'].values[0]#@
        nCodCC = get_nCodCC(Num_CC, OMIE_APP_KEY, OMIE_APP_SECRET)#@
    else:
        nCodCC = get_nCodCC('69024-5', OMIE_APP_KEY, OMIE_APP_SECRET)
    UF_cliente = get_cliente_UF(empresa, CNPJ)
    _, retencoes = get_impostos_materiais(empresa, CNPJ, UF_cliente, operacao, base_de_calculo=0)
    emails = consult_emails(empresa, CNPJ)#@
    if 'VENDA' in operacao:
        dizeres = generate_dizeres_PV(cabecalho, dizeres_main, cond_pgmt, vencimento, dados_bancarios, retencoes)
    elif 'REMESSA' in operacao:
        if type(dizeres_main.OBS) == np.float64 or type(dizeres_main.OBS) == float:
            if math.isnan(dizeres_main.OBS):
                dizeres=''
            else:
                dizeres=dizeres_main.OBS
        else:
            dizeres=dizeres_main.OBS
    ########
    endpoint = 'https://app.omie.com.br/api/v1/produtos/pedido/'
    payload = {}
    param = {}
    if vencimento != 'CONFIRMING':
        if type(vencimento) == str:
            vencimento = datetime.strptime(vencimento, '%d/%m/%Y').strftime('%d/%m/%Y')
        else:
            vencimento = vencimento.strftime('%d/%m/%Y')
        cod_parcela = "999"
        param['lista_parcelas'] = {
                                    "parcela": [{
                                                "data_vencimento": vencimento,
                                                "numero_parcela": 1,
                                                "percentual": 100,
                                                "valor":valor, #vtotal
                                                },]
                                  }
    else:
        cod_parcela = '999'

    payload['call'] = 'IncluirPedido'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    param['cabecalho'] = {
                            "codigo_cliente": codigo_cliente,
                            "codigo_pedido_integracao": str(int(cCodIntPV)),
                            "data_previsao": datetime.now().strftime('%d/%m/%Y'),
                            "etapa": "50",
                            "codigo_parcela": cod_parcela,
                          }
    param['det'] = []
    for material in materiais.iterrows():
        cfop = get_cfop(operacao, UF_cliente, material[1]['ST'])
        impostos, _ = get_impostos_materiais(empresa, CNPJ, UF_cliente, operacao, base_de_calculo=material[1]['V_TOTAL'])
        det = {
              "ide": {"codigo_item_integracao": int(material[0])+1, "simples_nacional":"S",},
              "produto": {
                            "cfop": cfop,
                            "codigo_produto": str(int(material[1]['COD'])),
                            "quantidade": float(material[1]['QTD']),
                            "valor_unitario": float(material[1]['V_UNITARIO'])
                          },
              "imposto":impostos,
              }
        param['det'].append(det)
    if type(dizeres_main['PROPOSTA']) == float or type(dizeres_main['PROPOSTA']) == np.float64:
        if math.isnan(dizeres_main['PROPOSTA']):
            proposta=''
        else:
            proposta = dizeres_main['PROPOSTA']
    else:
        proposta = dizeres_main['PROPOSTA']
    if type(dizeres_main['PEDIDO']) == float or type(dizeres_main['PEDIDO']) == np.float64:
        if math.isnan(dizeres_main['PEDIDO']):
            pedido=''
        else:
            pedido = dizeres_main['PEDIDO']
    else:
        pedido = dizeres_main['PEDIDO']
    param['informacoes_adicionais'] = {
                                        "codigo_categoria": "1.01.03",
                                        "codigo_conta_corrente": nCodCC,
                                        "numero_pedido_cliente":pedido, #pedido
                                        "numero_contrato":proposta, #proposta
                                        "dados_adicionais_nf":dizeres, #dizeres
                                        "consumidor_final": "S",
                                        "enviar_email": "S",
                                        "utilizar_emails":emails, #emails
                                        "tipo_documento":'NFE',
                                        'outros_detalhes':{'cNatOperacaoOd':natopr}
                                      },

    payload['param'] = [param,]
    response = requests.post(endpoint, json=payload).json()
    cCodIntPV_df = cCodIntPV_df.append({'ID':cCodIntPV}, ignore_index=True)
    cCodIntPV_df.to_excel(f'./INPUT/ID/IDpedido_de_venda{empresa.upper()}.xlsx', sheet_name=empresa.upper(), index=False)
    return response

def alterar_pv(codigo_pedido, cabecalho, dizeres_main, cond_pgmt, vencimento, materiais, valor):

    empresa = cabecalho.EMPRESA#@
    CNPJ = cabecalho.CNPJ#@
    operacao = dizeres_main.OPERACAO
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)#@
    codigo_cliente = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)#@
    dados_bancarios = consult_banco(empresa, CNPJ)#@
    Num_CC = dados_bancarios['CC'].values[0]#@
    nCodCC = get_nCodCC(Num_CC, OMIE_APP_KEY, OMIE_APP_SECRET)#@
    # CONTINUAR PELOS DIZERES
    UF_cliente = get_cliente_UF(empresa, CNPJ)
    _, retencoes = get_impostos_materiais(empresa, CNPJ, UF_cliente, operacao, base_de_calculo=0)
    dizeres = generate_dizeres_PV(cabecalho, dizeres_main, cond_pgmt, vencimento, dados_bancarios, retencoes)
    emails = consult_emails(empresa, CNPJ)#@
    ########

    endpoint = 'https://app.omie.com.br/api/v1/produtos/pedido/'
    payload = {}
    param = {}
    if vencimento != 'CONFIRMING':
        cod_parcela = "999"
        param['lista_parcelas'] = {
                                    "parcela": [{
                                                "data_vencimento": vencimento.strftime('%d/%m/%Y'),
                                                "numero_parcela": 1,
                                                "percentual": 100,
                                                "valor":valor, #vtotal
                                                },]
                                  }
    else:
        cod_parcela = '999'
    payload['call'] = 'AlterarPedidoVenda'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    param['cabecalho'] = {
                            "codigo_cliente": codigo_cliente,
                            "codigo_pedido": codigo_pedido,
                            "data_previsao": datetime.now().strftime('%d/%m/%Y'),
                            "etapa": "50",
                            "codigo_parcela": cod_parcela,
                          }
    param['det'] = []
    id_material=0
    for material in materiais.iterrows():
        cfop = get_cfop(operacao, UF_cliente, material[1]['ST'])
        impostos, _ = get_impostos_materiais(empresa, CNPJ, UF_cliente, operacao, base_de_calculo=valor)
        det = {
              "ide": {"codigo_item_integracao": int(material[0])+1, "simples_nacional":"S",},
              "produto": {
                            "cfop": cfop,
                            "codigo_produto": str(int(material[1]['COD'])),
                            "quantidade": float(material[1]['QTD']),
                            "valor_unitario": float(material[1]['V_UNITARIO'])
                          },
              "imposto":impostos,
              }
        param['det'].append(det)
        id_material = int(material[0])
    if type(dizeres_main['PROPOSTA']) == float or type(dizeres_main['PROPOSTA']) == np.float64:
        if math.isnan(dizeres_main['PROPOSTA']):
            proposta=''
    else:
        proposta = dizeres_main['PROPOSTA']
    if type(dizeres_main['PEDIDO']) == float or type(dizeres_main['PEDIDO']) == np.float64:
        if math.isnan(dizeres_main['PEDIDO']):
            pedido=''
    else:
        pedido = dizeres_main['PEDIDO']
    param['informacoes_adicionais'] = {
                                        "codigo_categoria": "1.01.03",
                                        "codigo_conta_corrente": nCodCC,
                                        "dados_adicionais_nf":dizeres, #dizeres
                                        "consumidor_final": "S",
                                        "enviar_email": "S",
                                        "utilizar_emails":emails, #emails
                                        "tipo_documento":'NFE',
                                      },

    payload['param'] = [param,]
    response = requests.post(endpoint, json=payload).json()
    return response

def faturar_pv(codigo_pedido, OMIE_APP_KEY, OMIE_APP_SECRET):
    endpoint = 'https://app.omie.com.br/api/v1/produtos/pedidovendafat/'
    payload = {}
    payload['call'] = 'FaturarPedidoVenda'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                        "nCodPed": codigo_pedido,
                        }]
    response = requests.post(endpoint, json=payload)
    return response.json()

def write_NFE(empresa, codigo_pedido, cond_pgmt, vencimento, tipo, operacao, pedido, proposta, solicitante, data_solicitacao, banco):
    faturados = pd.read_excel(f'./OUTPUT/FATURADOS/NF_{empresa.upper()}.xlsx', sheet_name=f'NF - {empresa.upper()}')
    nfe = get_nfe(empresa, codigo_pedido)
    if type(vencimento) != str:
        vencimento.strftime('%d/%m/%Y')
    CNPJ = nfe['nfDestInt']['cnpj_cpf']
    sleep(1)
    fantasia = get_cliente(empresa, CNPJ)['nome_fantasia']
    ########
    dct = {
    'SOLICITANTE':solicitante,
    'DATA SOLICITACAO':data_solicitacao,
    'PEDIDO':pedido,
    'PROPOSTA/CONTRATO':proposta,
    'NF':int(nfe['ide']['nNF']),
    'EMISSÃO':nfe['ide']['dEmi'],####
    'CLIENTE':fantasia,
    'CNPJ':CNPJ,
    'VALOR':nfe['titulos'][0]['nValorTitulo'], #14
    'CONDIÇÃO DE PAGAMENTO':cond_pgmt,#######
    'VENCIMENTO':vencimento,#####
    'VALOR DO ICMS':nfe['total']['ICMSTot']['vICMS'],
    'VALOR DO PIS':nfe['total']['ICMSTot']['vPIS'],
    'VALOR DO COFINS':nfe['total']['ICMSTot']['vCOFINS'],
    'BANCO': banco,
    'TAXA ANTECIPAÇÃO': 1,
    'OPERAÇÃO':operacao,
    'SITUAÇÃO':'APROVADA',
    'TIPO':tipo}
    faturados = faturados.append(dct, ignore_index=True)
    faturados = faturados.drop_duplicates()
    faturados.to_excel(f'./OUTPUT/FATURADOS/NF_{empresa.upper()}.xlsx', index=False, sheet_name=f'NF - {empresa}')
    print('ANOTADO NO RELATORIO GLEDSON!')
    return

def write_MATERIAIS_error(materiais):
    erros = pd.read_excel('./OUTPUT/ERROS/MATERIAIS.xlsx', sheet_name=f'MATERIAIS')
    erros = erros.append(materiais, ignore_index=True)
    erros = erros.drop_duplicates()
    try:
        erros.to_excel('./OUTPUT/ERROS/MATERIAIS.xlsx', sheet_name=f'MATERIAIS', index=False)
    except:
        print('Permissão para salvar erros de MATERIAIS negada')
    return

def write_NF_error(cabecalho, dizeres_main, cond_pgmt, materiais, obs=''):
    erros = pd.read_excel('./OUTPUT/ERROS/NF.xlsx', sheet_name=f'NF')
    CNPJ = cabecalho['CNPJ']
    if CNPJ.isnumeric() and len(CNPJ)>11:
        cabecalho['CNPJ'] = f"{CNPJ[:2]}.{CNPJ[2:5]}.{CNPJ[5:8]}/{CNPJ[8:12]}-{CNPJ[12:]}"
    elif CNPJ.isnumeric():
        cabecalho['CNPJ'] = f"{CNPJ[:3]}.{CNPJ[3:6]}.{CNPJ[6:9]}-{CNPJ[9:]}" #CPF
    OBS = pd.Series({'OBS_ERRO':obs})
    row = pd.concat([cabecalho, dizeres_main, cond_pgmt, OBS])
    erros = erros.append(row, ignore_index=True)
    erros = erros.drop_duplicates()
    try:
        erros.to_excel('./OUTPUT/ERROS/NF.xlsx', sheet_name=f'NF', index=False)
    except:
        print('Permissão para salvar erros de NF negada')
    write_MATERIAIS_error(materiais)

################################################################################
def criar_os(cabecalho, dizeres_main, cond_pgmt, vencimento):
    """
USAGE:
cabecalhos, dizeres_main, vencimetnos = get_data_NFS()
for i in range(0, len(cabecalho)):
    criar_os(cabecalhos.loc[i, :], dizeres_main.loc[i, :], vencimetnos.loc[i, :])
    print(f'OS referente ao pedido {dizeres_main.loc[i, 'PEDIDO']} criada com sucesso!')
    """
    # VALIDATION
    TEST = """
empresa = 'SOLAR'
CNPJ = '11.137.051/0011-58'
cabecalhos, dizeres_main, cond_pgmts = get_data_NFS()
cabecalho, dizeres_main, cond_pgmt = cabecalhos.loc[4,:], dizeres_main.loc[4, :], cond_pgmts.loc[4, :]
if cond_pgmt.empty:
    cond_pgmt = consult_cond_pgmt(empresa, CNPJ)
else:
    cond_pgmt = cond_pgmnts.loc[4, :]
vencimento = generate_vencimento(cond_pgmt, cabecalho['DATA FATURAMENTO'])

dados_bancarios = consult_banco(empresa, CNPJ)
codigo_servico = '0'+str(int(cabecalho['CODIGO MUNICIPAL']))
impostos, retencoes = get_impostos_servico(empresa, codigo_servico)

dizeres = generate_dizeres(cabecalho, dizeres_main, cond_pgmnt, dados_bancarios, retencoes)
emails = consult_emails(empresa, CNPJ)
    """
    empresa = cabecalho.EMPRESA # empresa emissora
    CNPJ = cabecalho.CNPJ # CNPJ DA EMPRESSA RECEPTORA
    if empresa != 'SOLAR' and empresa!= 'CONTRUAR':
        print('EMPRESA INVÁLIDA!!!')
        return


    codigo_servico = cabecalho['CODIGO MUNICIPAL']
    if type(codigo_servico) == np.float64 or type(codigo_servico) == np.float32:
        codigo_servico = '0'+str(int(codigo_servico))


    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    nCodCli = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)
    dados_bancarios = consult_banco(empresa, CNPJ)
    Num_CC = dados_bancarios['CC'].values[0]
    nCodCC = get_nCodCC(Num_CC, OMIE_APP_KEY, OMIE_APP_SECRET)
    impostos, retencoes = get_impostos_servico(empresa, codigo_servico)
    impostos, retencoes = get_impostos_particularidades(CNPJ, impostos, retencoes)
    dizeres = generate_dizeres(cabecalho, dizeres_main, cond_pgmt, vencimento, dados_bancarios, retencoes)
    emails = consult_emails(empresa, CNPJ)
    validation = [OMIE_APP_KEY, nCodCli, nCodCC, emails]
    cCodIntOS_df = pd.read_excel(f'./INPUT/ID/cCodIntOS{empresa.upper()}.xlsx', sheet_name=empresa.upper())
    cCodIntOS = cCodIntOS_df.loc[len(cCodIntOS_df)-1, 'ID']+1
    if None in validation:
        print(f'DADO INVÁLIDO')
        print(f"""
EMPRESA:{empresa} <=> API: {OMIE_APP_KEY}
CNPJ:{CNPJ} <=> COD:{nCodCli}
CC:{Num_CC} <=> COD:{nCodCC}
EMAILS:{emails}
DADOS BANCARIOS: {dados_bancarios}
VENCIMENTO: {vencimento}
IMPOSTOS_JSON: {impostos}
RETENCOES_DF: {retencoes}
        """)
        return
    # 'https://app.omie.com.br/api/v1/servicos/os/'
    # OPERATION
    endpoint = 'https://app.omie.com.br/api/v1/servicos/os/'
    payload = {}
    param = {}
    payload['call'] = 'IncluirOS'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    param['Cabecalho'] = {
                          "cCodIntOS": int(cCodIntOS),
                          "cCodParc": "000", # a vista
                          "nQtdeParc": 1,
                          "cEtapa": "50", #Faturar
                          "dDtPrevisao": datetime.now().strftime('%d/%m/%Y'),
                          "nCodCli": nCodCli,#########################################
                          }
    param['Departamentos'] = []
    param['Email'] = {
              "cEnvBoleto": "N",
              "cEnvLink": "S",
              "cEnviarPara": str(emails).replace(" '',", '').replace("'", '')
              }
    if type(dizeres_main['PROPOSTA']) == float or type(dizeres_main['PROPOSTA']) == np.float64:
        if math.isnan(dizeres_main['PROPOSTA']):
            proposta=''
        else:
            proposta = dizeres_main['PROPOSTA']
    else:
        proposta = dizeres_main['PROPOSTA']
    if type(dizeres_main['PEDIDO']) == float or type(dizeres_main['PEDIDO']) == np.float64:
        if math.isnan(dizeres_main['PEDIDO']):
            pedido=''
        else:
            pedido = dizeres_main['PEDIDO']
    else:
        pedido = dizeres_main['PEDIDO']
    param['InformacoesAdicionais'] = {
                              "cCodCateg": "1.01.02", # 'Clientes - Serviços Prestados'
                              "cNumPedido": pedido, #N_pedido
                              "cNumContrato": proposta, #N_contrato
                              "cDadosAdicNF": dizeres, # dizeres
                              "nCodCC": nCodCC, # codigo da conta corrente
                              }
    param['ServicosPrestados'] =  [{
                            "cCodServLC116": "14.06",
                            "cCodServMun": codigo_servico,
                            "cDescServ": cabecalho['DESCRICAO SERVICO'],
                            "cRetemISS": "N",
                            "cTribServ": "01", #Tributado no Municipio
                            "impostos": impostos,
                            "nQtde": 1,
                            "nValUnit": cabecalho['VALOR'],
                          },]
    payload['param'] = [param,]
    response = requests.post(endpoint, json=payload).json()
    cCodIntOS_df = cCodIntOS_df.append({'ID':cCodIntOS}, ignore_index=True)
    cCodIntOS_df.to_excel(f'./INPUT/ID/cCodIntOS{empresa.upper()}.xlsx', sheet_name=empresa.upper(), index=False)
    return response

def alterar_os(numero_os, cabecalho, dizeres_main, cond_pgmt, vencimento):
        """
USAGE:
cabecalhos, dizeres_main, cond_pgmts = get_data_NFS()
for i in range(0, len(cabecalho)):
    numero_os = get_numero_os(dizeres_main.loc[i, 'PEDIDO'], dizeres_main.loc[i, 'PROPOSTA'])
    alterar_os(numero_os, cabecalhos.loc[i, :], dizeres_main.loc[i, :], cond_pgmts.loc[i, :])
    print(f'OS referente ao pedido {dizeres_main.loc[i, 'PEDIDO']} criada com sucesso!')
        """
        # VALIDATION
        TEST = """
empresa = 'SOLAR'
CNPJ = '11.137.051/0011-58'
cabecalhos, dizeres_main, cond_pgmts = get_data_NFS()
cabecalho, dizeres_main, cond_pgmt = cabecalhos.loc[4,:], dizeres_main.loc[4, :], cond_pgmts.loc[4, :]
cond_pgmt = consult_cond_pgmt(empresa, CNPJ)
vencimento = generate_vencimento(cond_pgmt, cabecalhos.loc[4, 'DATA FATURAMENTO'])

dados_bancarios = consult_banco(empresa, CNPJ)

codigo_servico = '0'+str(int(cabecalho['CODIGO MUNICIPAL']))
impostos, retencoes = get_impostos_servico(empresa, codigo_servico)

dizeres = generate_dizeres(cabecalho, dizeres_main, vencimento, dados_bancarios, retencoes)
emails = consult_emails(empresa, CNPJ)
        """
        empresa = cabecalho.EMPRESA # empresa emissora
        CNPJ = cabecalho.CNPJ # CNPJ DA EMPRESSA RECEPTORA
        if empresa != 'SOLAR' and empresa!= 'CONTRUAR':
            print('EMPRESA INVÁLIDA!!!')
            return
        codigo_servico = cabecalho['CODIGO MUNICIPAL']
        if type(codigo_servico) == np.float64 or type(codigo_servico) == np.float32:
            codigo_servico = '0'+str(int(codigo_servico))
        OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
        nCodCli = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)
        dados_bancarios = consult_banco(empresa, CNPJ)
        Num_CC = dados_bancarios['CC'].values[0]
        nCodCC = get_nCodCC(Num_CC, OMIE_APP_KEY, OMIE_APP_SECRET)
        impostos, retencoes = get_impostos_servico(empresa, codigo_servico)
        impostos, retencoes = get_impostos_particularidades(CNPJ, impostos, retencoes)
        dizeres = generate_dizeres(cabecalho, dizeres_main, cond_pgmt, vencimento, dados_bancarios, retencoes)
        emails = consult_emails(empresa, CNPJ)
        validation = [OMIE_APP_KEY, nCodCli, nCodCC, emails]
        ###########
        if None in validation:
            print(f'DADO INVÁLIDO')
            print(f"""
EMPRESA:{empresa} <=> API: {OMIE_APP_KEY}
CNPJ:{CNPJ} <=> COD:{nCodCli}
CC:{Num_CC} <=> COD:{nCodCC}
EMAILS:{emails}
DADOS BANCARIOS: {dados_bancarios}
VENCIMENTO: {vencimento}
IMPOSTOS_JSON: {impostos}
RETENCOES_DF: {retencoes}
            """)
            return
        # 'https://app.omie.com.br/api/v1/servicos/os/'
        # OPERATION
        endpoint = 'https://app.omie.com.br/api/v1/servicos/os/'
        payload = {}
        param = {}
        payload['call'] = 'AlterarOS'
        payload['app_key'] = OMIE_APP_KEY
        payload['app_secret'] = OMIE_APP_SECRET
        param['Cabecalho'] = {
                              "nCodOS": numero_os,
                              "cCodParc": "000", # a vista
                              "nQtdeParc": 1,
                              "cEtapa": "50", #Faturar
                              "dDtPrevisao": datetime.now().strftime('%d/%m/%Y'),
                              "nCodCli": nCodCli,
                              }
        param['Departamentos'] = []
        param['Email'] = {
                  "cEnvBoleto": "N",
                  "cEnvLink": "S",
                  "cEnviarPara": str(emails).replace(" '',", '').replace("'", '')
                  }
        param['InformacoesAdicionais'] = {
                                  "cCodCateg": "1.01.02", # 'Clientes - Serviços Prestados'
                                  "cNumPedido": dizeres_main['PEDIDO'], #N_pedido
                                  "cNumContrato": dizeres_main['PROPOSTA'], #N_contrato
                                  "cDadosAdicNF": dizeres, # dizeres
                                  "nCodCC": nCodCC, # codigo da conta corrente
                                  }
        param['ServicosPrestados'] =  [{
                                "cCodServLC116": "14.06",
                                "cCodServMun": codigo_servico,
                                "cDescServ": cabecalho['DESCRICAO SERVICO'],
                                "cRetemISS": "N",
                                "cTribServ": "01", #Tributado no Municipio
                                "impostos": impostos,
                                "nQtde": 1,
                                "nValUnit": cabecalho['VALOR'],
                                'nSeqItem':1,
                              },]
        payload['param'] = [param,]
        response = requests.post(endpoint, json=payload).json()
        return response

def faturar_os(nCodOS, OMIE_APP_KEY, OMIE_APP_SECRET):
    endpoint = 'https://app.omie.com.br/api/v1/servicos/osp/'
    payload = {}
    payload['call'] = 'FaturarOS'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                        "nCodOS": nCodOS,
                        }]
    response = requests.post(endpoint, json=payload)
    return response.json()

def write_NFSE(empresa, nCodOS, cond_pgmt, vencimento, tipo, solicitante, data_solicitacao, banco, link):
    faturados = pd.read_excel(f'./OUTPUT/FATURADOS/NFS_{empresa.upper()}.xlsx', sheet_name=f'NFS - {empresa.upper()}')
    nfse = get_nfse(empresa, nCodOS)
    if 'CANCELAMENTO' in nfse.keys():
        situacao = 'CANCELADA'
    else:
        situacao = 'APROVADA'
    if 'cNumeroContrato' in nfse['OrdemServico'].keys():
        num_contrato = nfse['OrdemServico']['cNumeroContrato']
    else:
        num_contrato= ''
    if 'cPedidoCliente' in nfse['OrdemServico'].keys():
        num_pedido = nfse['OrdemServico']['cPedidoCliente']
    else:
        num_pedido= ''

    if type(vencimento) != str:
        vencimento.strftime('%d/%m/%Y')
    if 'cCNPJDestinatario' in nfse['Cabecalho']:
        CNPJ = nfse['Cabecalho']['cCNPJDestinatario']
    else:
        CNPJ = nfse['Cabecalho']['cCPFDestinatario']
    if '11137051' in CNPJ.replace('.', ''):
        fantasia = 'BOTICARIO'
    else:
        fantasia = get_cliente(empresa, CNPJ)['nome_fantasia']
    dct = {
    'SOLICITANTE':solicitante,
    'DATA SOLICITACAO':data_solicitacao,
    'PEDIDO':num_pedido,
    'PROPOSTA/CONTRATO':num_contrato,
    'NFS':nfse['Cabecalho']['nNumeroNFSe'],
    'EMISSÃO':nfse['Emissao']['cDataEmissao'],####
    'CLIENTE':fantasia,
    'CNPJ':CNPJ,
    'VALOR':nfse['Valores']['nValorTotalServicos'],
    'CONDIÇÃO DE PAGAMENTO':cond_pgmt,#######
    'VENCIMENTO':vencimento,#####
    'INSS (ALIQUOTA)':nfse['ListaServicos'][0]['nAliquotaINSS']/100,
    'PIS (ALIQUOTA)':nfse['ListaServicos'][0]['nAliquotaPIS']/100,
    'COFINS (ALIQUOTA)':nfse['ListaServicos'][0]['nAliquotaCOFINS']/100,
    'CSLL (ALIQUOTA)':nfse['ListaServicos'][0]['nAliquotaCSLL']/100,
    'IRRF (ALIQUOTA)':nfse['ListaServicos'][0]['nAliquotaIR']/100,
    'RETENÇÕES TOTAIS (ALIQUOTA)':(nfse['ListaServicos'][0]['nAliquotaINSS']+nfse['ListaServicos'][0]['nAliquotaPIS']+nfse['ListaServicos'][0]['nAliquotaCOFINS']+nfse['ListaServicos'][0]['nAliquotaCSLL']+nfse['ListaServicos'][0]['nAliquotaIR'])/100,
    'INSS (VALOR)':(nfse['ListaServicos'][0]['nAliquotaINSS']/100)*nfse['Valores']['nValorTotalServicos'],
    'PIS (VALOR)':(nfse['ListaServicos'][0]['nAliquotaPIS']/100)*nfse['Valores']['nValorTotalServicos'],
    'COFINS (VALOR)':(nfse['ListaServicos'][0]['nAliquotaCOFINS']/100)*nfse['Valores']['nValorTotalServicos'],
    'CSLL (VALOR)':(nfse['ListaServicos'][0]['nAliquotaCSLL']/100)*nfse['Valores']['nValorTotalServicos'],
    'IRRF (VALOR)':(nfse['ListaServicos'][0]['nAliquotaIR']/100)*nfse['Valores']['nValorTotalServicos'],
    'RETENÇÕES TOTAIS (VALOR)':((nfse['ListaServicos'][0]['nAliquotaINSS']/100)*nfse['Valores']['nValorTotalServicos'])+((nfse['ListaServicos'][0]['nAliquotaCSLL']/100)*nfse['Valores']['nValorTotalServicos'])+((nfse['ListaServicos'][0]['nAliquotaCOFINS']/100)*nfse['Valores']['nValorTotalServicos'])+((nfse['ListaServicos'][0]['nAliquotaPIS']/100)*nfse['Valores']['nValorTotalServicos'])+((nfse['ListaServicos'][0]['nAliquotaIR']/100)*nfse['Valores']['nValorTotalServicos']),
    'VALOR LIQUIDO':nfse['Valores']['nValorLiquido'],
    'BANCO': banco,
    'ANTECIPAÇÃO':np.nan,
    'TAXA ANTECIPAÇÃO':1,
    'RECEBIMENTO':np.nan,#
    'VALOR RECEBIDO':np.nan,
    'TIPO':tipo,
    'SITUAÇÃO':situacao,
    'nCodOS':nCodOS,
    'LINK':link,}
    faturados = faturados.append(dct, ignore_index=True)
    faturados = faturados.drop_duplicates()
    faturados.to_excel(f'./OUTPUT/FATURADOS/NFS_{empresa.upper()}.xlsx', index=False, sheet_name=f'NFS - {empresa}')
    print('ANOTADO NO RELATORIO GLEDSON!')
    return

def write_NFS_error(cabecalho, dizeres_main, cond_pgmt, obs=''):
    erros = pd.read_excel('./OUTPUT/ERROS/NFS.xlsx', sheet_name=f'NFS')
    CNPJ = cabecalho['CNPJ']
    if CNPJ.isnumeric() and len(CNPJ)>11:
        cabecalho['CNPJ'] = f"{CNPJ[:2]}.{CNPJ[2:5]}.{CNPJ[5:8]}/{CNPJ[8:12]}-{CNPJ[12:]}"
    elif CNPJ.isnumeric():
        cabecalho['CNPJ'] = f"{CNPJ[:3]}.{CNPJ[3:6]}.{CNPJ[6:9]}-{CNPJ[9:]}" #CPF
    OBS = pd.Series({'OBS_ERRO':obs})
    row = pd.concat([cabecalho, dizeres_main, cond_pgmt, OBS])
    erros = erros.append(row, ignore_index=True)
    erros = erros.drop_duplicates()
    try:
        erros.to_excel('./OUTPUT/ERROS/NFS.xlsx', sheet_name=f'NFS', index=False)
    except:
        print('Permissão para salvar erros de serviços negada')

def cancelar_NFSE(empresa, NFSe):
    faturamentos = pd.read_excel(f'./OUTPUT/FATURADOS/NFS_{empresa.upper()}.xlsx', sheet_name=f'NFS - {empresa.upper()}')
    faturamento = faturamentos.loc[faturamentos.loc[:, 'NFS']==NFSe, :]
    if faturamento['SITUAÇÃO'].values[0] == 'APROVADA':
        OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
        endpoint = 'https://app.omie.com.br/api/v1/servicos/osp/'
        payload = {}
        payload['call'] = 'CancelarOS'
        payload['app_key'] = OMIE_APP_KEY
        payload['app_secret'] = OMIE_APP_SECRET
        payload['param'] = [{
                              "nCodOS": int(faturamento['nCodOS'].values[0]),
                            }]
        response = requests.post(endpoint, json=payload)
        faturamentos.loc[faturamentos.loc[:, 'NFS']==NFSe, 'SITUAÇÃO'] = 'CANCELADA'
        faturamentos.to_excel(f'./OUTPUT/FATURADOS/NFS_{empresa.upper()}.xlsx', index=False, sheet_name=f'NFS - {empresa}')
    return
################################################################################
def alterar_pv_fila(cabecalho, dizeres_main, cond_pgmt, vencimento, lista_pv_faturar):
    _="""
    1 - CONSULTAR FILA DE FATURAMENTO
    """
    empresa = cabecalho.EMPRESA
    pv_proposta = dizeres_main.PROPOSTA
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    #lista_pv_faturar = listar_pv_faturar(empresa)
    vencimento = vencimento
    payload = {}
    endpoint = 'https://app.omie.com.br/api/v1/produtos/pedido/'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    for pv in lista_pv_faturar:
        if int(pv_proposta) == int(pv['cabecalho']['numero_pedido']):
            ####################################################################
            # EXCLUIR ITEM
            param = {}
            payload['call'] = 'AlterarPedidoVenda'
            if vencimento != 'CONFIRMING':
                cod_parcela = "999"
                param['lista_parcelas'] = {
                                            "parcela": [{
                                                        "data_vencimento": vencimento.strftime('%d/%m/%Y'),
                                                        "numero_parcela": 1,
                                                        "percentual": 100,
                                                        "valor":valor, #vtotal
                                                        },]
                                          }
            else:
                cod_parcela = '999'
            nCodCli = pv['cabecalho']['codigo_cliente']
            cliente = find_client_by(nCodCli, empresa)
            CNPJ = cliente['cnpj_cpf'].replace('.', '').replace('/', '').replace('-', '')
            UF_cliente = get_cliente_UF(empresa, CNPJ)
            param['cabecalho'] = {
                                    "codigo_cliente": nCodCli,
                                    "codigo_pedido": pv['cabecalho']['codigo_pedido'],
                                    "data_previsao": datetime.now().strftime('%d/%m/%Y'),
                                    "etapa": "50",
                                    "codigo_parcela": cod_parcela,
                                  }
            param['det'] = pv['det'].copy()
            cfop = get_cfop('PEDIDO DE VENDA', UF_cliente, ST=0)
            for i in range(len(param['det'])):
                #excluir
                param['det'][i]['ide']['acao_item'] = 'E'
                param['det'][i]['ide']["codigo_item_integracao"] = i+1
                #incluir
                impostos, _ = get_impostos_materiais(empresa, CNPJ, UF_cliente, 'PEDIDO DE VENDA', base_de_calculo=pv['det'][i]['produto']['valor_total'])
                param['det'].append(pv['det'][i].copy())
                param['det'][-1]['ide']['acao_item'] = ''
                param['det'][-1]['produto']['cfop'] = cfop
                param['det'][-1]['imposto'] = impostos
                param['det'][-1]['ide']["codigo_item_integracao"] = len(param['det'])
            payload['param'] = [param,]
            print('Produtos Removidos')
            ####################################################################
            # EMAILS, DIZERES, RETENCOES
            dados_bancarios = consult_banco(empresa, CNPJ)#@
            Num_CC = dados_bancarios['CC'].values[0]#@
            nCodCC = get_nCodCC(Num_CC, OMIE_APP_KEY, OMIE_APP_SECRET)#@
            _, retencoes = get_impostos_materiais(empresa, CNPJ, UF_cliente, 'PEDIDO DE VENDA', base_de_calculo=0)
            proposta = str(int(dizeres_main.PROPOSTA))
            pedido = str(dizeres_main.PEDIDO)
            dizeres = generate_dizeres_PV(cabecalho, dizeres_main, cond_pgmt, vencimento, dados_bancarios, retencoes)
            emails = consult_emails(empresa, CNPJ)
            param['informacoes_adicionais'] = {
                                                "codigo_categoria": "1.01.03",
                                                "codigo_conta_corrente": nCodCC,
                                                "dados_adicionais_nf":dizeres, #dizeres
                                                "consumidor_final": "S",
                                                "enviar_email": "S",
                                                "utilizar_emails":emails, #emails
                                                "tipo_documento":'NFE',
                                                "numero_pedido_cliente":pedido if len(pedido)<=30 else pedido[:30],
                                                "numero_contrato":proposta if len(proposta)<=30 else proposta[:30],
                                              },

            response = requests.post(endpoint, json=payload).json()
            pprint(response)
            print('Produtos Inserido')
            return response
    return None

### CANCELAR
def relatorio_claudenir_servico(mes_n):
    empresas = ['SOLAR', 'CONTRUAR']
    ORANGE = Color(rgb='FFf79646')
    meses = ['JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
    mes = meses[int(mes_n)-1]
    wb = Workbook()
    for empresa in empresas:
        faturamentos = pd.read_excel(f'./OUTPUT/FATURADOS/NFS_{empresa.upper()}.xlsx', sheet_name=f'NFS - {empresa.upper()}')
        ws = wb.create_sheet(f'{empresa} - SERVICOS')
        ############################################################################
        # HEADER
        thick = Side(style='thick')
        thin = Side(style='thin')
        ws.merge_cells('A1:H1')
        ws['A1'] = f'{empresa} - {mes} - SERVIÇOS'
        cells_header_month = [ws['A1'], ws['B1'], ws['C1'], ws['D1'], ws['E1'], ws['F1'], ws['G1'], ws['H1'], ws['I1']]
        for c in cells_header_month:                  #  00000000
            c.font = Font(b=True)
            c.alignment = Alignment(vertical='center', horizontal='center')
            c.fill = PatternFill(patternType='solid', fgColor=ORANGE)
            c.border = Border(top=thick, bottom=thick)

        ws['A1'].border = Border(left=thick, top=thick, bottom=thick)
        ws['I1'].border = Border(right=thick, top=thick, bottom=thick)
        ws['A2'], ws['B2'], ws['C2'], ws['D2'], ws['E2'], ws['F2'], ws['G2'], ws['H2'], ws['I2'] = 'N° NFS', 'EMISSÃO', 'CLIENTE', 'CNPJ/CPF', 'VALOR', 'VENCIMENTO', 'VALOR LÍQUIDO', 'RECEBIMETNO','TIPO'
        cells_header = [ws['A2'], ws['B2'], ws['C2'], ws['D2'], ws['E2'], ws['F2'], ws['G2'], ws['H2'], ws['I2']]
        for c in cells_header:                  #  00000000
            c.font = Font(b=True)
            c.alignment = Alignment(vertical='center', horizontal='center')
            c.fill = PatternFill(patternType='solid', fgColor=ORANGE)
            c.border = Border(left=thick, right=thick, top=thick, bottom=thick)
        ############################################################################
        # DATA
        faturamento = faturamentos.loc[faturamentos.loc[:, 'SITUAÇÃO']=='APROVADA', :]
        faturamento.loc[:, 'MES'] = faturamento.loc[:, 'EMISSÃO'].copy().apply(lambda x: str(x).split('/')[1] if '/' in str(x) else str(x).split('-')[1])
        data0 = faturamento.loc[faturamento.loc[:, 'MES']==f'{mes_n}', ['NFS', 'EMISSÃO', 'CLIENTE', 'CNPJ', 'VALOR', 'VENCIMENTO', 'VALOR LIQUIDO', 'RECEBIMENTO', 'TIPO']].reset_index()
        i=2
        for row in data0.iterrows():
            i=row[0]+3
            ws[f'A{i}'] = row[1]['NFS']
            ws[f'B{i}'] = row[1]['EMISSÃO']
            ws[f'B{i}'].number_format = 'DD/MM/YYYY'
            ws[f'C{i}'] = str(row[1]['CLIENTE']).replace('&amp;', '&')
            ws[f'D{i}'] = row[1]['CNPJ']
            ws[f'E{i}'] = row[1]['VALOR']
            ws[f'E{i}'].number_format = '"R$"#,##0.00'
            ws[f'F{i}'] = row[1]['VENCIMENTO']
            ws[f'F{i}'].number_format = 'DD/MM/YYYY'
            ws[f'G{i}'] = row[1]['VALOR LIQUIDO']
            ws[f'G{i}'].number_format = '"R$"#,##0.00'
            ws[f'H{i}'] = row[1]['RECEBIMENTO']
            ws[f'H{i}'].number_format = 'DD/MM/YYYY'
            ws[f'I{i}'] = row[1]['TIPO']
            cells_of_the_row = [ws[f'A{i}'], ws[f'B{i}'], ws[f'C{i}'], ws[f'D{i}'], ws[f'E{i}'], ws[f'F{i}'], ws[f'G{i}'], ws[f'H{i}'], ws[f'I{i}']]
            for c in cells_of_the_row:
                c.font = Font(b=False)
                c.alignment = Alignment(vertical='center', horizontal='center')
                c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        ########################################################################
        # FOOTER
        i = i+1
        ws.merge_cells(f'A{i}:D{i}')
        ws[f'A{i}'] = 'TOTAL'
        ws[f'E{i}'] = f'=SUM(E3:E{i-1})'
        ws[f'E{i}'].number_format = '"R$"#,##0.00'
        ws[f'G{i}'] = f'=SUM(G3:G{i-1})'
        ws[f'G{i}'].number_format = '"R$"#,##0.00'
        cells_footer_total = [ws[f'A{i}'], ws[f'B{i}'], ws[f'C{i}'], ws[f'D{i}']]
        for c in cells_footer_total:
            c.font = Font(b=True)
            c.alignment = Alignment(vertical='center', horizontal='center')
            c.fill = PatternFill(patternType='solid', fgColor=ORANGE)
            c.border = Border(top=thick, bottom=thick)
        ws[f'A{i}'].border = Border(left=thick, top=thick, bottom=thick)
        cells_footer = [ws[f'E{i}'], ws[f'F{i}'], ws[f'G{i}'], ws[f'H{i}'], ws[f'I{i}']]
        for c in cells_footer:                  #  00000000
            c.font = Font(b=True)
            c.alignment = Alignment(vertical='center', horizontal='center')
            c.fill = PatternFill(patternType='solid', fgColor=ORANGE)
            c.border = Border(left=thick, right=thick, top=thick, bottom=thick)
        ########################################################################
        # AUTO AJUSTE DAS COLUNAS
        dims = {}
        for row in ws.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
        for col, value in dims.items():
            ws.column_dimensions[col].width = value
    ############################################################################
    wb.remove(wb['Sheet'])
    wb.save(f'./OUTPUT/RELATORIOS/{mes_n} - RELATORIO DE FATURAMENTO {mes}.xlsx')
    return

def relatorio_claudenir_material(mes_n):
    empresas = ['SOLAR', 'CONTRUAR']
    ORANGE = Color(rgb='FFf79646')
    meses = ['JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
    mes = meses[int(mes_n)-1]
    wb = load_workbook(f'./OUTPUT/RELATORIOS/{mes_n} - RELATORIO DE FATURAMENTO {mes}.xlsx')
    for empresa in empresas:
        faturamentos = pd.read_excel(f'./OUTPUT/FATURADOS/NF_{empresa.upper()}.xlsx', sheet_name=f'NF - {empresa.upper()}')
        ws = wb.create_sheet(f'{empresa} - MATERIAIS')
        ############################################################################
        # HEADER
        thick = Side(style='thick')
        thin = Side(style='thin')
        ws.merge_cells('A1:G1')
        ws['A1'] = f'{empresa} - {mes} - MATERIAIS'
        cells_header_month = [ws['A1'], ws['B1'], ws['C1'], ws['D1'], ws['E1'], ws['F1'], ws['G1'], ws[f'H1']]
        for c in cells_header_month:                  #  00000000
            c.font = Font(b=True)
            c.alignment = Alignment(vertical='center', horizontal='center')
            c.fill = PatternFill(patternType='solid', fgColor=ORANGE)
            c.border = Border(top=thick, bottom=thick)

        ws['A1'].border = Border(left=thick, top=thick, bottom=thick)
        ws['G1'].border = Border(right=thick, top=thick, bottom=thick)
        ws['A2'], ws['B2'], ws['C2'], ws['D2'], ws['E2'], ws['F2'], ws['G2'], ws[f'H2'] = 'N° NF', 'EMISSÃO', 'CLIENTE', 'CNPJ/CPF', 'VALOR', 'VENCIMENTO', 'RECEBIMENTO', 'TIPO'
        cells_header = [ws['A2'], ws['B2'], ws['C2'], ws['D2'], ws['E2'], ws['F2'], ws['G2'], ws[f'H2']]
        for c in cells_header:                  #  00000000
            c.font = Font(b=True)
            c.alignment = Alignment(vertical='center', horizontal='center')
            c.fill = PatternFill(patternType='solid', fgColor=ORANGE)
            c.border = Border(left=thick, right=thick, top=thick, bottom=thick)
        ############################################################################
        # DATA
        faturamento = faturamentos.loc[faturamentos.loc[:, 'SITUAÇÃO'] != 'CANCELADA', :]
        faturamento = faturamento.loc[faturamento.loc[:, 'SITUAÇÃO'] != 'INUTILIZADA', :]
        faturamento = faturamento.loc[faturamento.loc[:, 'OPERAÇÃO'] == 'PEDIDO DE VENDA', :]
        faturamento.loc[:, 'MES'] = faturamento.loc[:, 'EMISSÃO'].copy().apply(lambda x: str(x).split('/')[1] if '/' in str(x) else str(x).split('-')[1])
        data0 = faturamento.loc[faturamento.loc[:, 'MES']==f'{mes_n}', ['NF', 'EMISSÃO', 'CLIENTE', 'CNPJ', 'VALOR', 'VENCIMENTO', 'RECEBIMENTO','TIPO']].reset_index()
        i = 3
        for row in data0.iterrows():
            i=row[0]+3
            ws[f'A{i}'] = row[1]['NF']
            ws[f'B{i}'] = row[1]['EMISSÃO']
            ws[f'B{i}'].number_format = 'DD/MM/YYYY'
            ws[f'C{i}'] = row[1]['CLIENTE'].replace('&amp;', '&')
            ws[f'D{i}'] = row[1]['CNPJ']
            ws[f'E{i}'] = row[1]['VALOR']
            ws[f'E{i}'].number_format = '"R$"#,##0.00'
            ws[f'F{i}'] = row[1]['VENCIMENTO']
            ws[f'F{i}'].number_format = 'DD/MM/YYYY'
            ws[f'G{i}'] = row[1]['RECEBIMENTO']
            ws[f'G{i}'].number_format = 'DD/MM/YYYY'
            ws[f'H{i}'] = row[1]['TIPO']
            cells_of_the_row = [ws[f'A{i}'], ws[f'B{i}'], ws[f'C{i}'], ws[f'D{i}'], ws[f'E{i}'], ws[f'F{i}'], ws[f'G{i}'], ws[f'H{i}']]
            for c in cells_of_the_row:
                c.font = Font(b=False)
                c.alignment = Alignment(vertical='center', horizontal='center')
                c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        ########################################################################
        # FOOTER
        i = i+1
        ws.merge_cells(f'A{i}:D{i}')
        ws[f'A{i}'] = 'TOTAL'
        ws[f'E{i}'] = f'=SUM(E3:E{i-1})'
        ws[f'E{i}'].number_format = '"R$"#,##0.00'
        cells_footer_total = [ws[f'A{i}'], ws[f'B{i}'], ws[f'C{i}'], ws[f'D{i}']]
        for c in cells_footer_total:
            c.font = Font(b=True)
            c.alignment = Alignment(vertical='center', horizontal='center')
            c.fill = PatternFill(patternType='solid', fgColor=ORANGE)
            c.border = Border(top=thick, bottom=thick)
        ws[f'A{i}'].border = Border(left=thick, top=thick, bottom=thick)
        cells_footer = [ws[f'E{i}'], ws[f'F{i}'], ws[f'G{i}'], ws[f'H{i}']]
        for c in cells_footer:                  #  00000000
            c.font = Font(b=True)
            c.alignment = Alignment(vertical='center', horizontal='center')
            c.fill = PatternFill(patternType='solid', fgColor=ORANGE)
            c.border = Border(left=thick, right=thick, top=thick, bottom=thick)
        ########################################################################
        # AUTO AJUSTE DAS COLUNAS
        dims = {}
        for row in ws.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
        for col, value in dims.items():
            ws.column_dimensions[col].width = value
    wb.save(f'./OUTPUT/RELATORIOS/{mes_n} - RELATORIO DE FATURAMENTO {mes}.xlsx')
    return
                        #'02'

def relatorio_claudenir(mes_n):
    relatorio_claudenir_servico(mes_n)
    relatorio_claudenir_material(mes_n)
    return

def relatorio_gledison(empresa):
    return

################################################################################
def get_NF_entrada(empresa, nota, fornecedor_CNPJ):
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    endpoint = 'https://app.omie.com.br/api/v1/contador/xml/'
    payload = {}
    payload['call'] = 'ListarDocumentos'
    payload['app_key'] = OMIE_APP_KEY
    payload['app_secret'] = OMIE_APP_SECRET
    payload['param'] = [{
                          "nPagina": 1,
                          "nRegPorPagina": 50,
                          "cModelo": "55",
                          'cOperacao':'0',
                          "dEmiInicial": "01/03/2022",
                          "dEmiFinal": "31/12/2022",
                        }]
    response = requests.post(endpoint, json=payload).json()
    pgs = response['nTotPaginas']
    for pg in range(pg, 0, -1):
        payload['param'][0]['nPagin']
        response = requests.post(endpoint, json=payload).json()
        registros = response['documentosEncontrados']
        for registro in registros:
            if registro['nNumero'] == nota:
                return registro['cXml']


def devolver_materiais(empresa, nota, fornecedor_CNPJ, materiais_id):
    xml_nf = get_NF_entrada(empresa, nota)
