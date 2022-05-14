from datetime import datetime, date
from sqlalchemy import Boolean, Column, ForeignKey, Integer, String, Float, DateTime, Date, PickleType, LargeBinary, SmallInteger
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy_utils import EmailType
from sqlalchemy.orm import relationship
from .database import Base

    

class PedidosCompra(Base):
    __tablename__='pedidoscompra'
    id = Column(Integer, primary_key=True, index=True)
    solicitante = Column(String)
    data_solicitacao = Column(Date)
    proposta = Column(Integer)
    cnpj_cpf_fornecedor = Column(Integer)
    tipo = Column(String)
    cond_pgmt = Column(String)
    itens = Column(PickleType)
    criacao = Column(Date, default=date.today())

class NotasFiscais(Base):
    __tablename__='notasfiscais'
    id = Column(Integer, primary_key=True, index=True)
    numero = Column(Integer, nullable=False)
    emissao = Column(Date)
    destinatario = Column(String)
    cnpj_cpf_d = Column(String)
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
    natureza_operacao = Column(String)
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
    solicitante = Column(String)
    data_solicitacao = Column(Date)
    execucao = Column(String)
    pedido = Column(String)
    proposta_contrato = Column(String)
    banco = Column(String)
    tipo = Column(String)
    status = Column(String)
    link = Column(String)
class NFM_Entradas(Base):
    __tablename__='nfm_entradas'
    id = Column(Integer, primary_key=True, index=True)
    notafiscalmateriais_id = Column(Integer, ForeignKey('notasfiscaismateriais.id'))
    solicitante = Column(String)
    data_solicitacao = Column(Date)
    pedido = Column(String)
    proposta_contrato = Column(String)
    banco = Column(String, nullable=True)
    tipo = Column(String)
    status = Column(String)
    classificacao = Column(String)
    xml = Column(LargeBinary)
class NFS_Saidas(Base):
    __tablename__='nfs_saidas'
    id = Column(Integer, primary_key=True, index=True)
    notafiscalservico_id = Column(Integer, ForeignKey('notasfiscaisservicos.id'))
    solicitante = Column(String)
    data_solicitacao = Column(Date)
    execucao = Column(String)
    pedido = Column(String)
    proposta_contrato = Column(String)
    banco = Column(String)
    tipo = Column(String)
    status = Column(String)
    link = Column(String)
class NFM_Saidas(Base):
    __tablename__='nfm_saidas'
    id = Column(Integer, primary_key=True, index=True)
    notafiscalmateriais_id = Column(Integer, ForeignKey('notasfiscaismateriais.id'))
    solicitante = Column(String)
    data_solicitacao = Column(Date)
    pedido = Column(String)
    proposta_contrato = Column(String)
    banco = Column(String, nullable=True)
    tipo = Column(String)
    status = Column(String)
    xml = Column(LargeBinary)
