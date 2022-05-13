from _database import schemas, crud, utils
from _database.database import Session
from datetime import date
from templates.email_template import Email
import tools


class Compras():
    def __init__(self):
        self.email_conn = Email('smtp.solarar.com.br', 'compras@solarar.com.br', 'Solar@2022')

    def gerar_pedido_compra(self, db: Session, pedido: schemas.PedidoCompra, obs: str|None = None):
        fornecedor_data = tools.get_fornecedor(pedido.cnpj_cpf_fornecedor)
        if not fornecedor_data:
            print('Fornecedor não cadastrado')
            return None
        db_pedido_compra = crud.create_pedido_compra(db=db, content=pedido)
        if not db_pedido_compra:
            print('Erro ao cadastrar pedido de compra')
            return None
        print('Pedido de Compra cadastrado com sucesso')
        tools.gerar_pdf_pedido_compra(
                pedido=db_pedido_compra.id, fornecedor='fornecedor',
                cnpj_cpf=db_pedido_compra.cnpj_cpf_fornecedor,
                endereco='endereco', itens=db_pedido_compra.itens, obs=obs)
        tools.enviar_email_pedido_compra(self.email_conn, pedido)
        return db_pedido_compra
    def enviar_pedido_compra(self, pedido_id):
        pass


if __name__ == '__main__':
    utils._create_database()
    global db
    db = utils.get_db()

def test_pedido_compra():
    pedido = schemas.PedidoCompra(
            empresa='SOLAR',
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
    compras = Compras()
    global db_pedido_compra
    db_pedido_compra = compras.gerar_pedido_compra(db=db, pedido=pedido)
