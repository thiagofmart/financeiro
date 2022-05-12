from datetime import datetime
from sqlalchemy import Boolean, Column, ForeignKey, Integer, String, Float, DateTime, Date, PickleType, LargeBinary
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy_utils import EmailType
from sqlalchemy.orm import relationship
from .database import Base, engine


class PedidoCompra(Base):
    __tablename__='pedidocompra'
    id = Column(Integer, primary_key=True, index=True)
    empresa = Column(String)
    solicitante = Column(String)
    data_solicitacao = Column(Date)
    proposta = Column(Integer)
    cnpj_cpf_fornecedor = Column(Integer)
    tipo = Column(String)
    itens = Column(JSONB)

class NotasFiscais(Base):
    __tablename__='notasfiscais'
    id = Column(Integer, primary_key=True, index=True)
    remetente = Column(String)
    cnpj_cpf_r = Column(String)
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

class NotasFiscaisServicos(NotasFiscais):
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
class NotasFiscaisMateriais(NotasFiscais):
    __tablename__ = 'notasfiscaismateriais'
    id = Column(Integer, primary_key=True, index=True)
    notafiscal_id = Column(Integer, ForeignKey('notasfiscais.id'))
    codigo_servico = Column(Integer)
    natureza_operacao = Column(String)
    icms = Column(Float)
    pis = Column(Float)
    cofins = Column(Float)
    materiais = Column(JSONB)

    entradas = relationship('NFM_Entradas', backref='notasfiscaismateriais')
    saidas = relationship('NFM_Saidas', backref='notasfiscaismateriais')

class NFS_Entradas(NotasFiscaisServicos):
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
class NFM_Entradas(NotasFiscaisMateriais):
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
class NFS_Saidas(NotasFiscaisServicos):
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
class NFM_Saidas(NotasFiscaisMateriais):
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

Base.metadata.create_all(engine)