from pydantic import BaseModel
from typing import List
from datetime import date


class Item(BaseModel):
    descricao: str
    qtd: float
    unidade: str
    unitario: float
    total: float

class PedidoCompra(BaseModel):
    empresa: str
    solicitante: str
    data_solicitacao: date
    proposta: int
    cnpj_cpf_fornecedor: int
    tipo: str
    cond_pgmt: str
    itens: List[Item]
