from pydantic import BaseModel
from typing import List
from datetime import date


class Email(BaseModel):
    email: str
    senha: str
    assinatura: str

class Enterprise(BaseModel):
    id: str
    formal_name: str
    informal_name: str
    address: str
    number: str
    complement: str
    district: str
    city: str
    county: str
    postal_code: str
    country: str
    telefone: str
    site: str
    email: str
    email_compra: Email
    email_nf: str
    email_faturamento: Email
    email_financeiro: Email
    ie: str #optional
    im: str #always
    logo: str

class Provider(BaseModel):
    id: int
    status: str
    raiz: str
    nome: str  
    emails_emi: List[str]
    emails_dest: List[str]

class Item(BaseModel):
    descricao: str
    qtd: float
    unidade: str
    unitario: float
    total: float
### Budgets ####################################
class ItemBudget(Item):
    tipo: str

class BudgetBase(BaseModel):
    cnpj_cpf_client: str
    estimator: str
    description: str
    cetegory: str
    manpower: bool
    items: List[ItemOrcamento]
    validity: date
    payment_terms: str
    
class BudgetCreate(BudgetBase):
    budget_difference_income: float
    
class Budget(BudgetBase):
    id: int
    bdi: float
    created_at: date


### Purchases #################################
class PurchaseRequestBase(BaseModel):
    requester: str
    budget_id: int
    provider_id: int
    _type: str
    items: List[Item]
    payment_terms: str
    
class PurchaseRequest(PurchaseRequestBase):
    id: int

### Bill ######################################