import os
from dotenv import load_dotenv

basedir = os.path.abspath(os.path.dirname(__file__))
load_dotenv(os.path.join(basedir, '.env'))

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'você-nunca-vai-adivinhar'
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or 'sqlite:///' + os.path.join(basedir, 'app.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    UPLOAD_FOLDER = os.path.join(basedir, 'uploads')
    TEMPLATE_DOCX_DIR = os.path.join(basedir, 'templates_docx')
    TEMPLATE_DOCX = os.path.join(TEMPLATE_DOCX_DIR, 'proposta_os', "Proposta OS 'AnoOS'-'NúmeroOs' - 'TítuloOS'.docx")
    TEMPLATE_RELATORIO_ENTREGA_DOCX = os.path.join(
        TEMPLATE_DOCX_DIR,
        'relatorio_entrega',
        "Relatório de Entrega OS 'AnoOS'-'NúmeroOs' - 'AssuntoOS'.docx"
    )
    FONTE_DADOS_DIR = os.path.join(basedir, 'fonte_de_dados')
    UPLOAD_PROPOSTAS_FOLDER = os.path.join(UPLOAD_FOLDER, 'propostas')
    UPLOAD_HUS_FOLDER = os.path.join(UPLOAD_FOLDER, 'hus')
    UPLOAD_RELATORIOS_FOLDER = os.path.join(UPLOAD_FOLDER, 'relatorios')
