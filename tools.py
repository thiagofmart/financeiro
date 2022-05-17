import pdfkit
from templates.email_template import Email
from _database import schemas, crud, utils
import pandas as pd
import numpy as np
from datetime import date
from constants import PATH_ROOT
import os

def get_fornecedor(cnpj_cpf: int):
    cnpj_cpf0 = cnpj_cpf
    cnpj_cpf = cnpj_cpf_to_str(cnpj_cpf)
    df = pd.read_excel(r'\\solsrv1\finc$\FINANCEIRO\COMPRAS\fornecedores.xlsx', sheet_name='FORNECEDORES')
    selection = df.loc[np.logical_and(df.loc[:, 'STATUS']=='ATIVO', df.loc[:, 'CNPJ/CPF']==cnpj_cpf), :].reset_index()
    if len(selection)==0:
        print(f'Fornecedor com CNPJ/CPF: {cnpj_cpf} não encontrado')
        return None
    select = selection.loc[0, :]
    emails_emi = select['EMAILS EMITENTE'].split(',')
    emails_dest = select['EMAILS DESTINATÁRIOS'].split(',')
    fornecedor_data = schemas.Fornecedor(
                status=select['STATUS'], raiz=select['RAIZ'],
                nome=select['NOME'], cnpj_cpf=cnpj_cpf0,
                emails_emi=emails_emi, emails_dest=emails_dest,
                )
    return fornecedor_data

def cnpj_cpf_to_str(cnpj_cpf: int):
    cnpj_cpf = str(cnpj_cpf)
    if len(cnpj_cpf)==14:
        formated_cnpj_cpf = f'{cnpj_cpf[:2]}.{cnpj_cpf[2:5]}.{cnpj_cpf[5:8]}/{cnpj_cpf[8:12]}-{cnpj_cpf[12:14]}'
    elif len(cnpj_cpf)==11:
        formated_cnpj_cpf = f'{cnpj_cpf[:3]}.{cnpj_cpf[3:6]}.{cnpj_cpf[6:9]}-{cnpj_cpf[9:11]}'
    else:
        raise TypeError
    return formated_cnpj_cpf

def gerar_pdf_pedido_compra(db_pedido_compra: schemas.dbPedidoCompra, fornecedor: schemas.Fornecedor, empresa: schemas.Empresa, obs: str|None = None):
    formated_cnpj_cpf = cnpj_cpf_to_str(fornecedor.cnpj_cpf)
    with open('./templates/pedido.html', 'r', encoding="UTF-8") as f_pedido:
        html = f_pedido.read()
        if obs:
            obs = f'OBS: <obs>{obs}</obs>'
        else:
            obs = ''
        html_data = ''
        total = 0
        for item in db_pedido_compra.itens:
            html_data = html_data+f"""<tr>
                    <td>{item.descricao}</td>
                    <td>{item.qtd}</td>
                    <td>{item.unidade}</td>
                    <td>R${'{:.2f}'.format(item.unitario).replace('.', ',')}</td>
                    <td>R${'{:.2f}'.format(item.total).replace('.', ',')}</td>
                </tr>"""
            total+=item.total
    html = html.replace('{{pedido}}', str(db_pedido_compra.id)
        ).replace('{{fornecedor_fantasia}}', fornecedor.nome
        ).replace('{{cnpj_cpf_fornecedor}}', formated_cnpj_cpf
        ).replace('{{endereco_fornecedor}}', 'endereco temporary'
        ).replace('{{email_nf}}', empresa.email_nf
        ).replace('{{obs}}', obs
        ).replace('{{logo}}', empresa.logo
        ).replace('{{itens}}', html_data
        ).replace('{{total}}', '{:.2f}'.format(total).replace('.', ',')
        ).replace('{{empresa_first}}', empresa.razao_social.split()[0]
        ).replace('{{empresa}}', empresa.razao_social
        ).replace('{{IE}}', empresa.ie
        ).replace('{{IM}}', empresa.im
        ).replace('{{logradouro}}', f'{empresa.endereco}, {empresa.numero}'
        ).replace('{{cep}}', empresa.cep
        ).replace('{{cidade}}', empresa.cidade
        ).replace('{{uf}}', empresa.uf
        ).replace('{{telefone}}', empresa.telefone
        ).replace('{{empresa_email}}', empresa.email
        ).replace('{{site}}', (f'<a href="https://{empresa.site}">{empresa.site}</a>' if empresa.site!='' else '')
        )
        # Footer

    options = {
        'margin-top': '0cm',
        'margin-right': '0cm',
        'margin-bottom': '0cm',
        'margin-left': '0cm',
        'encoding': "UTF-8",
        }
    config = pdfkit.configuration(wkhtmltopdf=r'.\venv_financeiro\wkhtmltopdf.exe')
    try:
        pdfkit.from_string(html, f'./output/{empresa.razao_social.split()[0]}/{empresa.razao_social.split()[0][0]}{db_pedido_compra.id}.pdf',configuration=config, options=options, css=f'./templates/{empresa.razao_social.split()[0]}/pdf_PC.css')
    except OSError as e:
        print(e)
    print('Arquivo gerado!')

def enviar_email(email_conn, receivers, subject, html_body, path_files):
    message = email_conn.create_message(receivers, subject, html_body, path_files)
    email_conn.send_email(message)

def enviar_email_pedido_compra(email_conn: Email, empresa: schemas.Empresa, db_pedido_compra: schemas.dbPedidoCompra, fornecedor: schemas.Fornecedor, obs=None):
    formated_cnpj_cpf = cnpj_cpf_to_str(fornecedor.cnpj_cpf)
    with (open('./templates/email_PC.html', 'r', encoding="UTF-8") as f_html, open(f'./templates/{empresa.razao_social.split()[0]}/email_PC.css', 'r', encoding="UTF-8") as f_css):
        html, css = f_html.read(), f_css.read()
        if obs:
            obs = f'OBS: <obs>{obs}</obs>'
        else:
            obs = ''
        html_data = ''
        total = 0
        for item in db_pedido_compra.itens:
            html_data = html_data+f"""<tr>
                    <td>{item.descricao}</td>
                    <td>{item.qtd}</td>
                    <td>{item.unidade}</td>
                    <td>R${'{:.2f}'.format(item.unitario).replace('.', ',')}</td>
                    <td>R${'{:.2f}'.format(item.total).replace('.', ',')}</td>
                </tr>"""
            total+=item.total
    html = html.replace('{{style}}', css
        ).replace('{{pedido}}', str(db_pedido_compra.id)
        ).replace('{{fornecedor_fantasia}}', fornecedor.nome
        ).replace('{{cnpj_cpf_fornecedor}}', formated_cnpj_cpf
        ).replace('{{endereco_fornecedor}}', 'endereco temporary'
        ).replace('{{email_nf}}', empresa.email_nf
        ).replace('{{obs}}', obs
        ).replace('{{logo}}', empresa.logo
        ).replace('{{itens}}', html_data
        ).replace('{{total}}', '{:.2f}'.format(total).replace('.', ',')
        ).replace('{{empresa_first}}', empresa.razao_social.split()[0]
        ).replace('{{empresa}}', empresa.razao_social
        ).replace('{{IE}}', empresa.ie
        ).replace('{{IM}}', empresa.im
        ).replace('{{logradouro}}', f'{empresa.endereco}, {empresa.numero}'
        ).replace('{{cep}}', empresa.cep
        ).replace('{{cidade}}', empresa.cidade
        ).replace('{{uf}}', empresa.uf
        ).replace('{{telefone}}', empresa.telefone
        ).replace('{{empresa_email}}', empresa.email
        ).replace('{{site}}', (f'<a href="https://{empresa.site}">{empresa.site}</a>' if empresa.site!='' else '')
        )
    enviar_email(
        email_conn=email_conn, receivers=fornecedor.emails_dest,
        subject=f'{empresa.fantasia} - PEDIDO DE COMPRA Nº{db_pedido_compra.id}',
        html_body=html, path_files=[os.path.join(PATH_ROOT, f'output/{empresa.razao_social.split()[0]}/{empresa.razao_social.split()[0][0]}{db_pedido_compra.id}.pdf.pdf'), ]
        )
    print('Email enviado!')
    return
