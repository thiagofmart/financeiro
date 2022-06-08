from datetime import datetime, date
from sqlalchemy import Boolean, Column, ForeignKey, Integer, String, Float, DateTime, Date, PickleType, LargeBinary, SmallInteger
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy_utils import EmailType
from sqlalchemy.orm import relationship
from .database import Base


class Enterprises(Base):
    __tablename__ = 'enterprises'
    id = Column(String, primary_key=True, index=True)
    informal_name = Column(String)
    formal_name = Column(String)
    address = Column(String)
    number = Column(String)
    complement = Column(String)
    district = Column(String)
    city = Column(String)
    county = Column(String)
    postal_code = Column(String)
    country = Column(String)
    phone = Column(String)
    site = Column(String)
    main_email = Column(String)
    purchase_email = Column(PickleType)
    bill_receiver_email = Column(PickleType)
    bill_sender_email = Column(PickleType)
    financial_email = Column(PickleType)
    additional_content = Column(PickleType)
    logo = Column(LargeBinary)


class AccreditedPersons(Base):
    __tablename__ = 'accreditedpersons'
    id = Column(String, primary_key=True)
    informal_name = Column(String)
    formal_name = Column(String)
    address = Column(String)
    number = Column(String)
    complement = Column(String)
    district = Column(String)
    city = Column(String)
    county = Column(String)
    postal_code = Column(String)
    country = Column(String)
    phone = Column(String)
    site = Column(String)
    main_email = Column(String)
    budget_email_receivers = Column(String)
    purchase_request_email_receiverss = Column(String)
    bill_email_receivers = Column(String)
    additional_content = Column(PickleType)

class Budgets(Base):
    __tablename__ = 'budgets'
    id = Column(Integer, primary_key=True, index=True)
    client_id = Column(Integer, ForeignKey('accreditedpersons.id'))
    estimator = Column(String)
    cetegory = Column(Integer)
    description = Column(String)
    manpower = Column(Boolean)
    items = Column(PickleType)
    validity = Column(Date)
    payment_terms = Column(String)
    bdi = Column(Float)
    created_at = Column(Date, default=date.today())
    

class PurchaseRequests(Base):
    __tablename__='purchaserequests'
    id = Column(Integer, primary_key=True, index=True)
    enterprise_id = Column(Integer, ForeignKey('enterprises.id'))
    requester = Column(String(200))
    budget_id = Column(Integer, ForeignKey('budgets.id'))
    provider_id = Column(Integer, ForeignKey('accreditedpersons.id'))
    _type = Column(String(200))
    payment_terms = Column(String(200))
    items = Column(PickleType)
    created_at = Column(Date, default=date.today())

class NotasFiscais(Base):
    __tablename__='notasfiscais'
    id = Column(Integer, primary_key=True, index=True)
    numero = Column(Integer, nullable=False)
    emissao = Column(Date)
    destinatario = Column(String(200))
    cnpj_cpf_d = Column(String(14))
    valor = Column(Float)
    qtd_parcelas = Column(Integer)
    n_parcela = Column(Integer)
    vencimento = Column(Date)

    notas_servicos = relationship('NotasFiscaisServicos', backref='notasfiscais')
    notas_materiais = relationship('NotasFiscaisMateriais', backref='notasfiscais')

class NotasFiscaisServicos(Base):
    __tablename__='notasfiscaisservicos'
    id = Column(Integer, primary_key=True, index=True)
    notafiscal_id = Column(Integer, ForeignKey('notasfiscais.id'))
    codigo_servico = Column(Integer)
    inss = Column(Float)
    irrf = Column(Float)
    csll = Column(Float)
    cofins = Column(Float)
    pis = Column(Float)
    liquido = Column(Float)

    entradas = relationship('NFS_Entradas', backref='notasfiscaisservicos')
    saidas = relationship('NFS_Saidas', backref='notasfiscaisservicos')
class NotasFiscaisMateriais(Base):
    __tablename__ = 'notasfiscaismateriais'
    id = Column(Integer, primary_key=True, index=True)
    notafiscal_id = Column(Integer, ForeignKey('notasfiscais.id'))
    codigo_servico = Column(Integer)
    natureza_operacao = Column(String(200))
    icms = Column(Float)
    pis = Column(Float)
    cofins = Column(Float)
    materiais = Column(PickleType)

    entradas = relationship('NFM_Entradas', backref='notasfiscaismateriais')
    saidas = relationship('NFM_Saidas', backref='notasfiscaismateriais')

class NFS_Entradas(Base):
    __tablename__='nfs_entradas'
    id = Column(Integer, primary_key=True, index=True)
    notafiscalservico_id = Column(Integer, ForeignKey('notasfiscaisservicos.id'))
    solicitante = Column(String(200))
    data_solicitacao = Column(Date)
    execucao = Column(String(200))
    pedido = Column(String(200))
    proposta_contrato = Column(String(200))
    banco = Column(String(200))
    tipo = Column(String(200))
    status = Column(String(200))
    link = Column(String(200))
class NFM_Entradas(Base):
    __tablename__='nfm_entradas'
    id = Column(Integer, primary_key=True, index=True)
    notafiscalmateriais_id = Column(Integer, ForeignKey('notasfiscaismateriais.id'))
    solicitante = Column(String(200))
    data_solicitacao = Column(Date)
    pedido = Column(String(200))
    proposta_contrato = Column(String(200))
    banco = Column(String(200), nullable=True)
    tipo = Column(String(200))
    status = Column(String(200))
    classificacao = Column(String(200))
    xml = Column(LargeBinary)
class NFS_Saidas(Base):
    __tablename__='nfs_saidas'
    id = Column(Integer, primary_key=True, index=True)
    notafiscalservico_id = Column(Integer, ForeignKey('notasfiscaisservicos.id'))
    solicitante = Column(String(200))
    data_solicitacao = Column(Date)
    execucao = Column(String(200))
    pedido = Column(String(200))
    proposta_contrato = Column(String(200))
    banco = Column(String(200))
    tipo = Column(String(200))
    status = Column(String(200))
    link = Column(String(200))
class NFM_Saidas(Base):
    __tablename__='nfm_saidas'
    id = Column(Integer, primary_key=True, index=True)
    notafiscalmateriais_id = Column(Integer, ForeignKey('notasfiscaismateriais.id'))
    solicitante = Column(String(200))
    data_solicitacao = Column(Date)
    pedido = Column(String(200))
    proposta_contrato = Column(String(200))
    banco = Column(String(200), nullable=True)
    tipo = Column(String(200))
    status = Column(String(200))
    xml = Column(LargeBinary)


