from _database import schemas, crud, utils
from sqlalchemy.orm import sessionmaker
from datetime import date
from templates.email_template import Email
import tools
from constants import empresas

class Compras():
    def __init__(self, empresa: str):
        self.empresa = empresas[empresa.upper()]
        self.db = utils.get_db(empresa.upper())
        self.email_conn = Email('smtp.solarar.com.br', 'compras@solarar.com.br', 'Solar@2022')

    def gerar_pedido_compra(self, pedido: schemas.PedidoCompra, obs: str|None = None):
        fornecedor_data = tools.get_fornecedor(pedido.cnpj_cpf_fornecedor)
        if not fornecedor_data:
            print('Fornecedor não cadastrado')
            return None
        db_pedido_compra = crud.create_pedido_compra(db=self.db, content=pedido)
        if not db_pedido_compra:
            print('Erro ao cadastrar pedido de compra')
            return None
        print('Pedido de Compra cadastrado com sucesso')
        tools.gerar_pdf_pedido_compra(
                db_pedido_compra=db_pedido_compra,
                fornecedor=fornecedor_data,
                empresa=self.empresa,
                )
        print('Pdf Gerado')
        return db_pedido_compra
    def enviar_pedido_compra(self, pedido_id, obs=None):
        db_pedido_compra = crud.get_pedido_compra_by_id(self.db, pedido_id)
        if not db_pedido_compra:
            print('Nº de pedido de compra inválido!')
            return None
        fornecedor_data = tools.get_fornecedor(db_pedido_compra.cnpj_cpf_fornecedor)
        if not fornecedor_data:
            print('Fornecedor não cadastrado')
            return None
        tools.enviar_email_pedido_compra(self.email_conn, self.empresa, db_pedido_compra, fornecedor_data, obs)
        print('E-mail enviado!')


if __name__ == '__main__':
    utils._create_database()

def test_pedido_compra(empresa):
    pedido = schemas.PedidoCompra(
            solicitante='BRUNO',
            data_solicitacao=date.today(),
            proposta=1,
            cnpj_cpf_fornecedor=11111111111111,
            tipo='MATERIAL',
            cond_pgmt='30 DDL',
            itens=[
                    schemas.Item(
                        descricao = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa 1',
                        qtd = 1,
                        unidade = 'UN',
                        unitario = 10.0,
                        total = 10.0,
                    ),
                    schemas.Item(
                        descricao = 'Item 2',
                        qtd = 2,
                        unidade = 'UN',
                        unitario = 20.0,
                        total = 20.0,
                    )
                    ]
            )
    compras = Compras(empresa)
    global db_pedido_compra
    db_pedido_compra = compras.gerar_pedido_compra(pedido=pedido)
