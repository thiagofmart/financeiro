from _database import schemas, crud, utils
from sqlalchemy.orm import sessionmaker
from datetime import date
from templates.email_template import Email
import tools
from constants import empresas
from fastapi import FastAPI, Depends, HTTPException

app = FastAPI()

@app.post("/api/v1/budget/create")
async def create_budget(paylaod: schemas.BudgetCreate):
    return {'none':'none'}
@app.post("api/v1/budget/send")
async def send_budget(payload: schemas.Budget):
    return
##############################################################

@app.post("/api/v1/purchase_request/create", response_model=schemas.PurchaseRequest)
async def create_purchase_request(purchaserequest: schemas.PurchaseRequestBase, db: Session=Depends(get_db_writer)):
    db_provider = crud.get_provider_by_id(purchaserequest.provider_id)
    if not db_provider:
        raise HTTPException(status_code=404, detail='provider not found')
    db_purchase_request = await crud.create_purchase_request(db=db, content=purchaserequest)
    if not db_purchase_request:
        raise HTTPException(status_code=404, detail='purchase request not found')
    ## APPROVED VALIDATION!
    return db_pedido_compra
@app.post("/api/v1/purchase_request/send")
async def send_purchase_request(purchase_request_id: int, db: Session=Depends(get_db_read)):
    db_pedido_compra = crud.get_purchase_request_by_id(db, purchase_request_id)
    if not db_pedido_compra:
        print('Nº de pedido de compra inválido!')
        return None
    fornecedor_data = tools.get_fornecedor(db_pedido_compra.cnpj_cpf_fornecedor)
    if not fornecedor_data:
        print('Fornecedor não cadastrado')
        return None
    tools.enviar_email_pedido_compra(self.email_conn, self.empresa, db_pedido_compra, fornecedor_data, obs)
    print('E-mail enviado!')
    return {'none':'none'}
@app.post("/api/v1/purchase/save_invoice")
async def save_invoice(invoice_xml):
    return {'none':'none'}

##############################################################
@app.post("/api/v1/operational/service/create_order")
async def create_service_order():
    return {'none':'none'}
@app.post("/api/v1/operational/service/execute")
async def execute_service():
    return {'none':'none'}
@app.post("/api/v1/operational/service/report")
async def report_service():
    return {'none':'none'}
###############################################################
@app.post("api/v1/bill/create")
async def create_bill():
    return {'none':'none'}
@app.post("api/v1/bill/send")
async def send_bill():
    return {'none':'none'}
###############################################################



class App():
    def __init__(self, empresa: str):
        self.empresa = empresas[empresa.upper()]
        self.db = utils.get_db(empresa.upper())
        self.orcamento = self.Orcamento()
        self.compra = self.Compra(self.empresa.email_compra, self.empresa, self.db)
        self.faturamento = self.Faturamento(self.empresa.email_faturamento)
        self.relatorio = self.Relatorios()
    
    class Compra():
        def __init__(self, email_conn: Email, empresa_data: schemas.Empresa, db: sessionmaker):
            self.empresa = empresa_data
            self.db = db
            self.email_conn = email_conn

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
    empresa = App(empresa)
    global db_pedido_compra
    db_pedido_compra = empresa.compra.gerar_pedido_compra(pedido=pedido)
