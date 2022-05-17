from pydantic import BaseModel
from typing import List
from datetime import date

_="""
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
compras = Compras('SOLAR')
db_pedido_compra = compras.gerar_pedido_compra(pedido=pedido)



"""
class Empresa(BaseModel):
    razao_social: str
    fantasia: str
    cnpj: str
    endereco: str
    numero: str
    complemento: str
    bairro: str
    cidade: str
    uf: str
    cep: str
    telefone: str
    site: str
    email: str
    email_nf: str
    ie: str #optional
    im: str #always
    logo: str

class Fornecedor(BaseModel):
    status: str
    raiz: str
    nome: str
    cnpj_cpf: int
    emails_emi: List[str]
    emails_dest: List[str]

class Item(BaseModel):
    descricao: str
    qtd: float
    unidade: str
    unitario: float
    total: float


class PedidoCompra(BaseModel):
    solicitante: str
    data_solicitacao: date
    proposta: int
    cnpj_cpf_fornecedor: int
    tipo: str
    cond_pgmt: str
    itens: List[Item]
class dbPedidoCompra(PedidoCompra):
    id: int
