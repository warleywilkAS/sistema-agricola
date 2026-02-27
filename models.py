from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class FormularioSoja(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Identificação
    numero_produtor = db.Column(db.String(100))
    regiao = db.Column(db.String(100))
    municipio = db.Column(db.String(100))
    meso_idr = db.Column(db.String(100))
    area_soja = db.Column(db.Float)
    produtividade_media = db.Column(db.Float)
    cultivar = db.Column(db.String(100))
    bt = db.Column(db.String(10))
    data_plantio = db.Column(db.String(20))
    data_emergencia = db.Column(db.String(20))
    houve_adversidade = db.Column(db.String(10))
    qual_adversidade = db.Column(db.String(100))
    nome_coletor = db.Column(db.String(200))
    unidade_municipal = db.Column(db.String(200))
    
    # Conhecimento MIP e MID
    conhecimento_mid = db.Column(db.String(10))
    utiliza_mid = db.Column(db.String(10))
    conhecimento_mip = db.Column(db.String(10))
    utiliza_mip = db.Column(db.String(10))
    
    # Controle Plantas Invasoras
    herbicida_dessecacao_alvo = db.Column(db.String(50))
    herbicida_dessecacao_aplicacoes = db.Column(db.Integer)
    herbicida_pre_alvo = db.Column(db.String(50))
    herbicida_pre_aplicacoes = db.Column(db.Integer)
    herbicida_pos_alvo = db.Column(db.String(50))
    herbicida_pos_aplicacoes = db.Column(db.Integer)
    
    # Tratamento sementes e outros
    tratamento_sementes = db.Column(db.String(10))
    sal_mistura = db.Column(db.String(10))
    controle_biologico = db.Column(db.String(10))
    
    # FBN
    inoculacao_sementes = db.Column(db.String(10))
    forma_inoculacao = db.Column(db.String(50))
    coinoculacao = db.Column(db.String(10))
    co_mo = db.Column(db.String(10))
    co_mo_aplicacao = db.Column(db.String(50))
    
    # Relacionamentos para pulverizações
    pulverizacoes = db.relationship('Pulverizacao', backref='formulario', lazy=True, cascade='all, delete-orphan')

class Pulverizacao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    formulario_id = db.Column(db.Integer, db.ForeignKey('formulario_soja.id'), nullable=False)
    tipo = db.Column(db.String(50))  # pre_plantio, pos_1, pos_2, etc.
    data = db.Column(db.String(20))
    classe_produto = db.Column(db.String(50))
    alvo = db.Column(db.String(200))