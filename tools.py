import pdfkit
from templates.email_template import Email
from _database import schemas, crud, utils
import pandas as pd
import numpy as np


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
                nome=select['NOME'], cnpj_cpf=cnpj0,
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

def gerar_pdf_pedido_compra(db_pedido_compra: schemas.dbPedidoCompra, fornecedor_data, obs: str|None = None):
    formated_cnpj_cpf = cnpj_cpf_to_str(cnpj_cpf)
    with open('./templates/pedido.html', 'r', encoding="UTF-8") as f_pedido:
        s_pedido = f_pedido.read()
        # HEADER
        if obs:
            obs = f'OBS: <obs>{obs}</obs>'
        else:
            obs = ''
        s_pedido = s_pedido.replace('{{pedido}}', str(db_pedido_compra.id)).replace('{{fornecedor}}', fornecedor).replace('{{cnpj_cpf}}', formated_cnpj_cpf).replace('{{endereco}}', endereco).replace('{{obs}}', obs)
        # DATA
        html_data = ''
        total = 0
        for item in db.pedido_compra.itens:
            html_data = html_data+f"""<tr>
                    <td>{item.descricao}</td>
                    <td>{item.qtd}</td>
                    <td>{item.unidade}</td>
                    <td>R${'{:.2f}'.format(item.unitario).replace('.', ',')}</td>
                    <td>R${'{:.2f}'.format(item.total).replace('.', ',')}</td>
                </tr>"""
            total+=item.total
        s_pedido = s_pedido.replace('{{itens}}', html_data).replace('{{total}}', '{:.2f}'.format(total).replace('.', ','))
        # Footer
        html=s_pedido

    path_wkhtmltopdf = r'.\venv_financeiro\wkhtmltopdf.exe'
    options = {
        'margin-top': '0cm',
        'margin-right': '0cm',
        'margin-bottom': '0cm',
        'margin-left': '0cm',
        'encoding': "UTF-8",
        }
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    try:
        pdfkit.from_string(html, f'./output/pc_{pedido}_{datetime.today}.pdf',configuration=config, options=options, css='./templates/style.css')
    except OSError:
        pass
    print('Arquivo gerado!')

def enviar_email(email_conn, receivers, subject, html_body, path_files):
    message = email_conn.create_message(receivers, subject, body, path_files)
    email_conn.send_email(message)

def enviar_email_pedido_compra(email_conn: Email, pedido_compra: schemas.dbPedidoCompra, fornecedor: schemas.Fornecedor):
    enviar_email(
        email_conn=email_conn, receivers=fornecedor.emails_dest,
        subject=f'SOLAR AR CONDICIONADO LTDA - PEDIDO DE COMPRA Nº{pedido_compra.id}',
        html_body=html_body, path_files=['./output/PC',]
        )
    print('Email enviado!')
    return
