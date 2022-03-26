from functions import *

#\\solsrv1\man$\CONTRATOS ATUANTES\TELHANORTE\CONTOLE GERAL - PROGRAMA\Programa TelhaNorte.xlsm
# CONTROLE DE PROPOSTAS


def faturar_pendencias_NF(cabecalho, dizeres_main, cond_pgmt, materiais):

    pedido, proposta, empresa, CNPJ = dizeres_main['PEDIDO'], dizeres_main['PROPOSTA'], cabecalho['EMPRESA'], cabecalho['CNPJ']
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    cond_pgmti = cond_pgmt.copy()
    cond_pgmt.dropna(how='all', inplace=True)
    dados_bancarios = consult_banco(empresa, CNPJ)
    banco = dados_bancarios.BANCO.values[0]
    if cond_pgmt.empty:
        cond_pgmt = consult_cond_pgmt(empresa, CNPJ)
        if len(cond_pgmt.values)==0:
            write_NFS_error(cabecalho, dizeres_main, cond_pgmti, obs='FALHA AO ENCONTRAR A CONDIÇÃO DE PAGAMENTO')
            return 'FALHA AO ENCONTRAR A CONDIÇÃO DE PAGAMENTO'
    else:
        cond_pgmt = pd.DataFrame(columns=cond_pgmti.keys(), data=[cond_pgmti.values,])
    vencimento = generate_vencimento(cond_pgmt, cabecalho['DATA FATURAMENTO'])
    tipo = dizeres_main.TIPO
    operacao = dizeres_main.OPERACAO
    if 'REMESSA' in operacao:
        banco = np.nan
    if tipo.upper() == 'PREVENTIVA':
        nCodCli = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)
        pv_cadastradas = consultar_pv(pedido, proposta, nCodCli, OMIE_APP_KEY, OMIE_APP_SECRET)
    else:
        pv_cadastradas = consultar_pv(pedido, proposta, None, OMIE_APP_KEY, OMIE_APP_SECRET)
    valor = materiais['V_TOTAL'].sum()
    if pv_cadastradas == None:
        response1 = criar_pv(cabecalho, dizeres_main, cond_pgmt, vencimento, materiais, valor)
        print(response1)
        print(f'PV referente ao pedido {dizeres_main["PEDIDO"]} cadastrada com sucesso!')
        #response 1 output =>
        #>>>{'codigo_pedido': 1729064251, 'codigo_pedido_integracao': '1003', 'codigo_status': '0', 'descricao_status': 'Pedido cadastrado com sucesso!', 'numero_pedido': '000000000000022'}
        codigo_pedido = response1['codigo_pedido']
        response2 = faturar_pv(codigo_pedido, OMIE_APP_KEY, OMIE_APP_SECRET)
        print(f'PV referente ao pedido {dizeres_main["PEDIDO"]} faturada com sucesso!')
        ##### cond pgmt
        cond_pgmt_=cond_pgmt.dropna(axis=1)
        if cond_pgmt_.columns[0] == 'DDL':
            cond_pgmt_ = f'{int(cond_pgmt_.values[0][0])} DDL'
        elif cond_pgmt_.columns[0] == 'FIXO':
            cond_pgmt_ = f'FIXO {cond_pgmt_.values[0][0]}'
        elif cond_pgmt_.columns[0] == 'CONFIRMING':
            cond_pgmt_ = 'CONFIRMING'
        ##### cond pgmt
        sleep(10)
        if type(dizeres_main["PEDIDO"]) == float or type(dizeres_main["PEDIDO"]) == np.float64:
            if math.isnan(dizeres_main["PEDIDO"]):
                pedido = ''
            else:
                pedido = dizeres_main["PEDIDO"]
        else:
            pedido = dizeres_main["PEDIDO"]
        if type(dizeres_main["PROPOSTA"]) == float or type(dizeres_main["PROPOSTA"]) == np.float64:
            if math.isnan(dizeres_main["PROPOSTA"]):
                proposta = ''
            else:
                proposta = dizeres_main["PROPOSTA"]
        else:
            proposta = dizeres_main["PROPOSTA"]
        write_NFE(empresa, codigo_pedido, cond_pgmt_, vencimento, tipo, operacao, pedido, proposta, cabecalho['SOLICITANTE'], cabecalho['DATA SOLICITACAO'], banco)
        return 'OS cadastrada e Faturada'
    elif len(pv_cadastradas['pendentes']) == 1:
        pv = pv_cadastradas['pendentes'][0]
        det_pv = pv['det']
        codigo_pedido = pv['cabecalho']['codigo_pedido']
        response1 = alterar_pv(codigo_pedido, cabecalho, dizeres_main, cond_pgmt, vencimento, materiais, valor, det_pv)
        print(response1)
        print(f'OS referente ao pedido { dizeres_main["PEDIDO"]} alterada com sucesso!')
        codigo_pedido = response1['codigo_pedido']
        response2 = faturar_pv(codigo_pedido, OMIE_APP_KEY, OMIE_APP_SECRET)
        print(response)
        print(f'OS referente ao pedido { dizeres_main["PEDIDO"]} faturada com sucesso!')
        ##### cond pgmt
        cond_pgmt_=cond_pgmt.dropna(axis=1)
        if cond_pgmt_.columns[0] == 'DDL':
            cond_pgmt_ = f'{int(cond_pgmt_.values[0][0])} DDL'
        elif cond_pgmt_.columns[0] == 'FIXO':
            cond_pgmt_ = f'FIXO {cond_pgmt_.values[0][0]}'
        elif cond_pgmt_.columns[0] == 'CONFIRMING':
            cond_pgmt_ = 'CONFIRMING'
        ##### cond pgmt
        sleep(10)
        if type(dizeres_main["PEDIDO"]) == float or type(dizeres_main["PEDIDO"]) == np.float64:
            if math.isnan(dizeres_main["PEDIDO"]):
                pedido = ''
            else:
                pedido = dizeres_main["PEDIDO"]
        else:
            pedido = dizeres_main["PEDIDO"]
        if type(dizeres_main["PROPOSTA"]) == float or type(dizeres_main["PROPOSTA"]) == np.float64:
            if math.isnan(dizeres_main["PROPOSTA"]):
                proposta = ''
            else:
                pedido = dizeres_main["PROPOSTA"]
        else:
            proposta = dizeres_main["PROPOSTA"]
        write_NFE(empresa, codigo_pedido, cond_pgmt_, vencimento, tipo, operacao, pedido, proposta, cabecalho['SOLICITANTE'], cabecalho['DATA SOLICITACAO'], banco)
        return 'OS alterada e Faturada'
    elif len(pv_cadastradas['pendentes']) > 1:
        return f'Há mais de uma PV referente ao pedido/proposta pendentes'
    elif len(pv_cadastradas['faturados']) >= 1:
        return f'PV referente ao pedido/proposta ja foi faturado!'
    elif len(pv_cadastradas['cancelados']) >= 1:
        response1 = criar_pv(cabecalho, dizeres_main, cond_pgmt, vencimento, materiais, valor)
        print(f'PV referente ao pedido {dizeres_main["PEDIDO"]} cadastrada com sucesso!')
        #response 1 output =>
        #>>>{'codigo_pedido': 1729064251, 'codigo_pedido_integracao': '1003', 'codigo_status': '0', 'descricao_status': 'Pedido cadastrado com sucesso!', 'numero_pedido': '000000000000022'}
        codigo_pedido = response1['codigo_pedido']
        response2 = faturar_pv(codigo_pedido, OMIE_APP_KEY, OMIE_APP_SECRET)
        print(f'PV referente ao pedido {dizeres_main["PEDIDO"]} faturada com sucesso!')
        ##### cond pgmt
        cond_pgmt_=cond_pgmt.dropna(axis=1)
        if cond_pgmt_.columns[0] == 'DDL':
            cond_pgmt_ = f'{int(cond_pgmt_.values[0][0])} DDL'
        elif cond_pgmt_.columns[0] == 'FIXO':
            cond_pgmt_ = f'FIXO {cond_pgmt_.values[0][0]}'
        elif cond_pgmt_.columns[0] == 'CONFIRMING':
            cond_pgmt_ = 'CONFIRMING'
        ##### cond pgmt
        sleep(5)
        if type(dizeres_main["PEDIDO"]) == float or type(dizeres_main["PEDIDO"]) == np.float64:
            if math.isnan(dizeres_main["PEDIDO"]):
                pedido = ''
            else:
                pedido = dizeres_main["PEDIDO"]
        else:
            pedido = dizeres_main["PEDIDO"]
        if type(dizeres_main["PROPOSTA"]) == float or type(dizeres_main["PROPOSTA"]) == np.float64:
            if math.isnan(dizeres_main["PROPOSTA"]):
                proposta = ''
            else:
                proposta = dizeres_main["PROPOSTA"]
        else:
            proposta = dizeres_main["PROPOSTA"]
        write_NFE(empresa, codigo_pedido, cond_pgmt_, vencimento, tipo, operacao, pedido, proposta, cabecalho['SOLICITANTE'], cabecalho['DATA SOLICITACAO'], banco)
        return 'OS cadastrada e Faturada'
    else:
        print('nada')
        return None

def faturar_pendencias_NFS(cabecalho, dizeres_main, cond_pgmt):

    pedido, proposta, empresa, CNPJ = dizeres_main['PEDIDO'], dizeres_main['PROPOSTA'], cabecalho['EMPRESA'], cabecalho['CNPJ']
    OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
    cond_pgmti = cond_pgmt.copy()
    dados_bancarios = consult_banco(empresa, CNPJ)
    banco = dados_bancarios.BANCO.values[0]
    cond_pgmt.dropna(how='all', inplace=True)
    if cond_pgmt.empty:
        cond_pgmt = consult_cond_pgmt(empresa, CNPJ)
        if len(cond_pgmt.values)==0:
            return 'FALHA AO ENCONTRAR A CONDIÇÃO DE PAGAMENTO'
    else:
        cond_pgmt = pd.DataFrame(columns=cond_pgmti.keys(), data=[cond_pgmti.values,])
    vencimento = generate_vencimento(cond_pgmt, cabecalho['DATA FATURAMENTO'])
    tipo = dizeres_main['TIPO']
    if tipo.upper() == 'PREVENTIVA':
        nCodCli = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)
        os_cadastradas = consultar_os(pedido, proposta, nCodCli, OMIE_APP_KEY, OMIE_APP_SECRET)
    else:
        os_cadastradas = consultar_os(pedido, proposta, None, OMIE_APP_KEY, OMIE_APP_SECRET)
    if os_cadastradas == None:
        response1 = criar_os(cabecalho, dizeres_main, cond_pgmt, vencimento)
        if 'faultstring' in response1.keys():
            print(response1['faultstring'])
        else:
            print(f'OS referente ao pedido {dizeres_main["PEDIDO"]} cadastrada com sucesso!')
        nCodOS = response1['nCodOS']
        response2 = faturar_os(nCodOS, OMIE_APP_KEY, OMIE_APP_SECRET)
        print(f'OS referente ao pedido {dizeres_main["PEDIDO"]} faturada com sucesso!')
        ##### cond pgmt
        cond_pgmt_=cond_pgmt.dropna(axis=1)
        if cond_pgmt_.columns[0] == 'DDL':
            cond_pgmt_ = f'{int(cond_pgmt_.values[0][0])} DDL'
        elif cond_pgmt_.columns[0] == 'FIXO':
            cond_pgmt_ = f'FIXO {cond_pgmt_.values[0][0]}'
        elif cond_pgmt_.columns[0] == 'CONFIRMING':
            cond_pgmt_ = f'CONFIRMING {cond_pgmt_.values[0][0]}'

        ##### cond pgmt
        sleep(5)
        status = get_Status_OS(nCodOS, empresa)
        link = status['ListaRpsNfse'][0]['danfe']
        write_NFSE(empresa, response2['nCodOS'], cond_pgmt_, vencimento, tipo, cabecalho['SOLICITANTE'], cabecalho['DATA SOLICITACAO'], banco, link)
        return 'OS cadastrada e Faturada'
    elif len(os_cadastradas['pendentes']) == 1:
        os = os_cadastradas['pendentes'][0]
        nCodOS = os['Cabecalho']['nCodOS']
        response1 = alterar_os(nCodOS, cabecalho, dizeres_main, cond_pgmt, vencimento)
        print(f'OS referente ao pedido { dizeres_main["PEDIDO"]} alterada com sucesso!')
        response2 = faturar_os(nCodOS, OMIE_APP_KEY, OMIE_APP_SECRET)
        print(f'OS referente ao pedido { dizeres_main["PEDIDO"]} faturada com sucesso!')
        ##### cond pgmt
        cond_pgmt_=cond_pgmt.dropna(axis=1)
        if cond_pgmt_.columns[0] == 'DDL':
            cond_pgmt_ = f'{int(cond_pgmt_.values[0][0])} DDL'
        elif cond_pgmt_.columns[0] == 'FIXO':
            cond_pgmt_ = f'FIXO {cond_pgmt_.values[0][0]}'
        elif cond_pgmt_.columns[0] == 'CONFIRMING':
            cond_pgmt_ = 'CONFIRMING'
        ##### cond pgmt
        sleep(5)
        status = get_Status_OS(nCodOS, empresa)
        link = status['ListaRpsNfse'][0]['danfe']
        write_NFSE(empresa, nCodOS, cond_pgmt_, vencimento, tipo, cabecalho['SOLICITANTE'], cabecalho['DATA SOLICITACAO'], banco, link)
        return 'OS alterada e Faturada'
    elif len(os_cadastradas['pendentes']) > 1:
        return f'Há mais de uma OS referente ao pedido/proposta pendentes'
    elif len(os_cadastradas['faturados']) >= 1:
        return f'OS referente ao pedido/proposta ja foi faturado!'
    elif len(os_cadastradas['cancelados']) >= 1:
        response1 = criar_os(cabecalho, dizeres_main, cond_pgmt, vencimento)
        print(f'OS referente ao pedido {dizeres_main["PEDIDO"]} cadastrada com sucesso!')
        nCodOS = response1['nCodOS']
        response2 = faturar_os(nCodOS, OMIE_APP_KEY, OMIE_APP_SECRET)
        print(f'OS referente ao pedido {dizeres_main["PEDIDO"]} faturada com sucesso!')
        ##### cond pgmt
        cond_pgmt_=cond_pgmt.dropna(axis=1)
        if cond_pgmt_.columns[0] == 'DDL':
            cond_pgmt_ = f'{int(cond_pgmt_.values[0][0])} DDL'
        elif cond_pgmt_.columns[0] == 'FIXO':
            cond_pgmt_ = f'FIXO {cond_pgmt_.values[0][0]}'
        elif cond_pgmt_.columns[0] == 'CONFIRMING':
            cond_pgmt_ = 'CONFIRMING'
        ##### cond pgmt
        sleep(5)
        status = get_Status_OS(nCodOS, empresa)
        link = status['ListaRpsNfse'][0]['danfe']
        write_NFSE(empresa, response2['nCodOS'], cond_pgmt_, vencimento, tipo, cabecalho['SOLICITANTE'], cabecalho['DATA SOLICITACAO'], banco, link)
        return 'OS cadastrada e Faturada'
    else:
        print('nada')
        return None

################################################################################

def iniciar_ciclo_faturamento_NF():
    """
    # STEPS
0- TENTAR LER OS DADOS NA PLANILHA DE INPUT ####check
1- PARA CADA LINHA EM NF-e SEPARAR TODOS OS MATERIAIS QUE EXISTEM EM MATERIAIS INPUT
2- VERIFICAR SE CADA MATERIAL JA EXISTE, CASO NEGATIVO CADASTRAR O MATERIAL
3- VERIFICAR SE O PV JA EXISTE, CASO POSITIVO ALTERAR O PV E CASO NEGATIVO CADASTRAR PV
4- FATURAR PV

TEST
cabecalhos, dizeres_main, cond_pgmts, datam = get_data_NF_root()
i=0
cabecalho, dizeres_m, cond_pgmt, empresa = cabecalhos.loc[i, :], dizeres_main.loc[i, :], cond_pgmts.loc[i, :], cabecalhos.loc[i, 'EMPRESA']
materiais = datam.loc[datam.loc[:, 'ID']==cabecalho['ID'], :].copy()
for material in materiais.iterrows():
    print(material[1]['DESCRICAO'])
    response = cadastrar_material(empresa, material[1])
    if response.__repr__()[-5:-2] == '500':
        print(response.json())
        print('ja existe')
        codigo = response.json()['faultstring'][77:-1]
        response = consultar_material(empresa, codigo)
        materiais.loc[material[0], 'COD'] = response['codigo_produto']
    else:
        print(response.json())
        print('não existe')
        materiais.loc[material[0], 'COD'] = response.json()['codigo_produto']

#INPUT status = faturar_pendencias_NF(cabecalho, dizeres_m, cond_pgmt, materiais)
################################################################################ start FATURAR_PENDENCIAS_NF
pedido, proposta, empresa, CNPJ = dizeres_m['PEDIDO'], dizeres_m['PROPOSTA'], cabecalho['EMPRESA'], cabecalho['CNPJ']
OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
cond_pgmti = cond_pgmt.copy()
cond_pgmt.dropna(how='all', inplace=True)
cond_pgmt = pd.DataFrame(columns=cond_pgmti.keys(), data=[cond_pgmti.values,])
vencimento = generate_vencimento(cond_pgmt, cabecalho['DATA FATURAMENTO'])
tipo = dizeres_m['TIPO']
operacao = dizeres_m['OPERACAO']
if tipo.upper() == 'PREVENTIVA':
    nCodCli = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)
    pv_cadastradas = consultar_pv(pedido, proposta, nCodCli, OMIE_APP_KEY, OMIE_APP_SECRET)
else:
    pv_cadastradas = consultar_pv(pedido, proposta, None, OMIE_APP_KEY, OMIE_APP_SECRET)


valor = materiais['V_TOTAL'].sum()

 criar_pv ou alterar_pv ##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@##

#"" SE TIVER PENDENCIAS
pv = pv_cadastradas['pendentes'][0]
det_pv = pv['det']
det_pv1 = []
for det in det_pv:
    det['produto']['quantidade'] = 0
    det_pv1.append(det)
codigo_pedido = pv['cabecalho']['codigo_pedido']
response1 = alterar_pv(codigo_pedido, cabecalho, dizeres_m, cond_pgmt, vencimento, materiais, valor, det_pv1)
""


#
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
UF_cliente = get_cliente_UF(empresa, CNPJ)
impostos, retencoes = get_impostos_materiais(empresa, CNPJ, UF_cliente)
dizeres = generate_dizeres_PV(cabecalho, dizeres_m, cond_pgmt, vencimento, dados_bancarios, retencoes)
emails = consult_emails(empresa, CNPJ)#@

        """
    try:
        cabecalhos, dizeres_main, cond_pgmts, datam = get_data_NF_root()
        sleep(1)
    except:
        print('FALHA NA COLETA DE DADOS!')
        sleep(30)
    for i in range(0, len(cabecalhos)):
        cabecalho, dizeres, cond_pgmt, empresa = cabecalhos.loc[i, :], dizeres_main.loc[i, :], cond_pgmts.loc[i, :], cabecalhos.loc[i, 'EMPRESA']
        print(f'PEDIDO:{dizeres["PEDIDO"]} - PROPOSTA:{dizeres["PROPOSTA"]}')
        materiais = datam.loc[datam.loc[:, 'ID']==cabecalho['ID'], :].copy()
        if type(dizeres['PEDIDO'])==float and type(dizeres['PROPOSTA'])==float:
            if math.isnan(dizeres['PEDIDO']) and math.isnan(dizeres['PROPOSTA']):
                write_NFS_error(cabecalho, dizeres, cond_pgmt, obs='NEM PEDIDO E NEM PROPOSTA ENCONTRADA')
                pass
        if len(materiais)==0:
            pass
            _="""
            OMIE_APP_KEY, OMIE_APP_SECRET = get_API(cabecalho['EMPRESA'])
            if dizeres['TIPO'].upper() == 'PREVENTIVA':
                nCodCli = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)
                pv_cadastradas = consultar_pv(dizeres['PEDIDO'], dizeres['PROPOSTA'], nCodCli, OMIE_APP_KEY, OMIE_APP_SECRET)
            else:
                pv_cadastradas = consultar_pv(dizeres['PEDIDO'], dizeres['PROPOSTA'], None, OMIE_APP_KEY, OMIE_APP_SECRET)
            pv = pv_cadastradas['pendentes'][0]
            det_pv = pv['det']
            for det in det_pv:
                dct = {
                'ID':cabecalho['ID'],
                'DESCRICAO':det['produto']['descricao'],
                'NCM':det['produto']['ncm'].replace('.', ''),
                'ST':0,
                'QTD':det['produto']['quantidade'],
                'UNIDADE':det['produto']['unidade'],
                'V_UNITARIO':det['produto']['valor_unitario'],
                'V_TOTAL':det['produto']['valor_total'],
                }
                row = pd.Series(dct)
                materiais = materiais.append(row, ignore_index=True)
            pv.dell
            """
        else:
            lst_cod_material = []
            for material in materiais.iterrows():
                print(material[1]['DESCRICAO'])
                response = cadastrar_material(empresa, material[1])
                if response.__repr__()[-5:-2] == '500':
                    if 'NCM não cadastrada' in response.json()['faultstring']:
                        print('NCM INVÁLIDO')
                    print(response.json())
                    #print('ja existe')
                    codigo = response.json()['faultstring'][77:-1]
                    response = consultar_material(empresa, codigo)
                    #print(response.__repr__())
                    materiais.loc[material[0], 'COD'] = response['codigo_produto']
                else:
                    print(response.json())
                    print('não existe')
                    materiais.loc[material[0], 'COD'] = response.json()['codigo_produto']
            status = faturar_pendencias_NF(cabecalho, dizeres, cond_pgmt, materiais)
            print(status)
    print('CICLO NFS FINALIZADO')
    return

def iniciar_ciclo_faturamento_NFS():
    """
TESTS
#STEP 1
cabecalhos, dizeres_main, cond_pgmnt = get_data_NFS_root()
i=0
cabecalho, dizeres, cond_pgmt = cabecalhos.loc[i, :], dizeres_main.loc[i, :], cond_pgmnt.loc[i, :]

#STEP 2
dizeres_main = dizeres
pedido, proposta, empresa, CNPJ = dizeres_main['PEDIDO'], dizeres_main['PROPOSTA'], cabecalho['EMPRESA'], cabecalho['CNPJ']
OMIE_APP_KEY, OMIE_APP_SECRET = get_API(empresa)
cond_pgmti = cond_pgmt.copy()
cond_pgmt.dropna(how='all', inplace=True)
cond_pgmt = pd.DataFrame(columns=cond_pgmti.keys(), data=[cond_pgmti.values,])
vencimento = generate_vencimento(cond_pgmt, cabecalho['DATA FATURAMENTO'])
tipo = dizeres_main['TIPO']
nCodCli = get_nCodCli(CNPJ, OMIE_APP_KEY, OMIE_APP_SECRET)
os_cadastradas = consultar_os(pedido, proposta, nCodCli, OMIE_APP_KEY, OMIE_APP_SECRET)

# STEP 3 - criar_os ##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@##@
empresa = cabecalho.EMPRESA # empresa emissora
CNPJ = cabecalho.CNPJ # CNPJ DA EMPRESSA RECEPTORA
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
# STEP 4 - DIZERES
    """
    try:
        cabecalhos, dizeres_main, cond_pgmnt = get_data_NFS_root()
        sleep(1)
    except:
        print('FALHA NA COLETA DE DADOS!')
        sleep(30)
    for i in range(0, len(cabecalhos)):
        sleep(0.5)
        cabecalho, dizeres, cond_pgmt = cabecalhos.loc[i, :], dizeres_main.loc[i, :], cond_pgmnt.loc[i, :]
        if type(dizeres['PEDIDO'])==float and type(dizeres['PROPOSTA'])==float:
            if math.isnan(dizeres['PEDIDO']) and math.isnan(dizeres['PROPOSTA']):
                write_NFS_error(cabecalho, dizeres, cond_pgmt, obs='NEM PEDIDO E NEM PROPOSTA ENCONTRADA')
                pass
        print(f'PEDIDO:{dizeres["PEDIDO"]} - PROPOSTA:{dizeres["PROPOSTA"]}')
        status = faturar_pendencias_NFS(cabecalho, dizeres, cond_pgmt)
        print(status)
    print('CICLO NFS FINALIZADO')

def iniciar_ciclo_faturamento_NF_fila():
    _="""

NÃO PRECISA PREENCHER CNPJ E MATERIAIS.
é necessario alterar a proposta na OS para linkar com a PLANILHA



TEST:
cabecalhos, dizeres_main, cond_pgmts, datam = get_data_NF_root()
i=0
cabecalho, dizeres, cond_pgmt = cabecalhos.loc[i, :], dizeres_main.loc[i, :], cond_pgmts.loc[i, :]
cond_pgmti = cond_pgmt.copy()
cond_pgmt.dropna(how='all', inplace=True)
if cond_pgmt.empty:
    cond_pgmt = consult_cond_pgmt(cabecalho.EMPRESA, cabecalho.CNPJ)
else:
    cond_pgmt = pd.DataFrame(columns=cond_pgmti.keys(), data=[cond_pgmti.values,])

vencimento = generate_vencimento(cond_pgmt, cabecalho['DATA FATURAMENTO'])
lista_pv_faturar = listar_pv_faturar(cabecalho.EMPRESA)
response = alterar_pv_fila(cabecalho, dizeres, cond_pgmt, vencimento, lista_pv_faturar)
41832.61
46028.36
"""
    cabecalhos, dizeres_main, cond_pgmts, datam = get_data_NF_root()
    print('DADOS COLETADOS')
    sleep(1)
    for i in range(0, len(cabecalhos)):
        cabecalho, dizeres, cond_pgmt = cabecalhos.loc[i, :], dizeres_main.loc[i, :], cond_pgmts.loc[i, :]
        OMIE_APP_KEY, OMIE_APP_SECRET = get_API(cabecalho.EMPRESA.upper())
        cond_pgmti = cond_pgmt.copy()
        cond_pgmt.dropna(how='all', inplace=True)
        if cond_pgmt.empty:
            cond_pgmt = consult_cond_pgmt(cabecalho.EMPRESA, cabecalho.CNPJ)
        else:
            cond_pgmt = pd.DataFrame(columns=cond_pgmti.keys(), data=[cond_pgmti.values,])
        vencimento = generate_vencimento(cond_pgmt, cabecalho['DATA FATURAMENTO'])
        lista_pv_faturar = listar_pv_faturar(cabecalho.EMPRESA)
        response = alterar_pv_fila(cabecalho, dizeres, cond_pgmt, vencimento, lista_pv_faturar)
        if response == None:
            print(f'proposta {dizeres.PROPOSTA} não encontrada')
        else:
            codigo_pedido = response['codigo_pedido']
            response = faturar_pv(codigo_pedido, OMIE_APP_KEY, OMIE_APP_SECRET)
            pprint(response)
            print('FATURADA')
            ##### cond pgmt
            cond_pgmt_=cond_pgmt.dropna(axis=1)
            if cond_pgmt_.columns[0] == 'DDL':
                cond_pgmt_ = f'{int(cond_pgmt_.values[0][0])} DDL'
            elif cond_pgmt_.columns[0] == 'FIXO':
                cond_pgmt_ = f'FIXO {cond_pgmt_.values[0][0]}'
            elif cond_pgmt_.columns[0] == 'CONFIRMING':
                cond_pgmt_ = 'CONFIRMING'
            ##### cond pgmt
            sleep(5)
            if type(dizeres.PEDIDO) == float or type(dizeres.PEDIDO) == np.float64:
                if math.isnan(dizeres.PEDIDO):
                    pedido = ''
                else:
                    pedido = dizeres.PEDIDO
            else:
                pedido = dizeres.PEDIDO
            if type(dizeres.PROPOSTA) == float or type(dizeres.PROPOSTA) == np.float64:
                if math.isnan(dizeres.PROPOSTA):
                    proposta = ''
                else:
                    proposta = dizeres.PROPOSTA
            else:
                proposta = dizeres_main["PROPOSTA"]
            write_NFE(cabecalho.EMPRESA, codigo_pedido, cond_pgmt_, vencimento, dizeres.TIPO, dizeres.OPERACAO, pedido, proposta, cabecalho['SOLICITANTE'], cabecalho['DATA SOLICITACAO'], banco=np.nan)
    print('CICLO NFS da fila FINALIZADO')
    return
