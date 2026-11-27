import os
import hmac
import time
import secrets
import calendar
import re
import shutil
import socket
import json
from datetime import datetime, date, timedelta
from io import BytesIO
from threading import Lock
from functools import wraps

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from sqlalchemy import event, text, inspect
from sqlalchemy.engine import Engine
from flask import (
    Flask,
    render_template_string,
    request,
    redirect,
    url_for,
    send_file,
    send_from_directory,
    flash,
    session,
    get_flashed_messages,
    jsonify,
    Response,
)
from flask_sqlalchemy import SQLAlchemy

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# ==============================================================================
# CONFIGURAÇÃO DO BANCO DE DADOS (PostgreSQL via DATABASE_URL ou SQLite Local)
# ==============================================================================
DB_PATH = os.path.join(BASE_DIR, "os.db")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
EVIDENCIAS_DIR = os.path.join(BASE_DIR, "uploads", "evidencias_os")
PCM_FILES_DIR = os.path.join(BASE_DIR, "uploads", "pcm_arquivos")
os.makedirs(BACKUP_DIR, exist_ok=True)
os.makedirs(EVIDENCIAS_DIR, exist_ok=True)
os.makedirs(PCM_FILES_DIR, exist_ok=True)

EXTENSOES_EVIDENCIA = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
}
EXTENSOES_PCM_EXCEL = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls": "application/vnd.ms-excel",
}
TAMANHO_MAX_EVIDENCIA = 6 * 1024 * 1024  # 6 MB por foto
TAMANHO_MAX_PCM_EXCEL = 25 * 1024 * 1024  # 25 MB por planilha Excel
MAX_FOTOS_EVIDENCIA_POR_OS = 12


def obter_database_uri():
    raw_url = (os.environ.get("DATABASE_URL") or "").strip()
    if raw_url:
        if raw_url.startswith("postgres://"):
            raw_url = raw_url.replace("postgres://", "postgresql+psycopg2://", 1)
        elif raw_url.startswith("postgresql://") and "+psycopg2" not in raw_url:
            raw_url = raw_url.replace("postgresql://", "postgresql+psycopg2://", 1)
        return raw_url, True
    return f"sqlite:///{DB_PATH}", False


app_db_uri, IS_POSTGRES = obter_database_uri()


def carregar_chave_secreta():
    env_secret = (os.environ.get("SECRET_KEY") or "").strip()
    if env_secret:
        return env_secret
    caminho = os.path.join(BASE_DIR, ".secret_key")
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            valor = f.read().strip()
            if valor:
                return valor
    except OSError:
        pass
    valor = secrets.token_hex(32)
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            f.write(valor)
    except OSError:
        pass
    return valor


app = Flask(__name__)
app.config["SECRET_KEY"] = carregar_chave_secreta()
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["PERMANENT_SESSION_LIFETIME"] = 8 * 3600
app.config["MAX_CONTENT_LENGTH"] = 48 * 1024 * 1024
app.config["SQLALCHEMY_DATABASE_URI"] = app_db_uri
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

if IS_POSTGRES:
    app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
        "pool_pre_ping": True,
        "pool_recycle": 300,
    }
else:
    app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
        "connect_args": {"timeout": 30},
    }

db = SQLAlchemy(app)


@event.listens_for(Engine, "connect")
def configurar_sqlite(dbapi_connection, connection_record):
    if IS_POSTGRES:
        return
    try:
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")
        cursor.execute("PRAGMA synchronous=NORMAL;")
        cursor.execute("PRAGMA foreign_keys=ON;")
        cursor.close()
    except Exception:
        pass


# ==============================================================================
# CONSTANTES E CONFIGURAÇÕES OPERACIONAIS
# ==============================================================================
FORMATO_DATA = "%d/%m/%Y %H:%M:%S"
PREFIXO_ANO_OS = "26"
NUMERO_OS_INICIAL = 1917

PREFIXO_ANO_SC = "26"
NUMERO_SC_INICIAL = 517

os_lock = Lock()
estoque_lock = Lock()
sc_lock = Lock()

SISTEMA_USUARIO = "manutencao"
SISTEMA_SENHA = "manutencao123"
MAX_TENTATIVAS_LOGIN = 5
BLOQUEIO_LOGIN_SEG = 300
tentativas_login = {}

TIPOS_MANUTENCAO = ("Preventiva", "Corretiva")
PRIORIDADES_VALIDAS = ("Baixa", "Média", "Alta", "Crítica")
STATUS_OS_VALIDOS = ("Aberta", "Aguardando Peça", "Finalizada")
STATUS_SC_VALIDOS = (
    "Aguardando Aprovação",
    "Em Cotação",
    "Aprovado",
    "Reprovado",
    "Finalizado / Comprado",
)
TIPOS_SC_VALIDOS = (
    "S.C. Emergencial",
    "S.C. Programada",
    "S.C. Reposição / Estoque",
)

JORNADA_REFERENCIA = 9.0
FUNCIONARIOS_EQUIPE = [
    {
        "valor": "Cesar",
        "nome": "Cesar",
        "funcao": "Mecânico",
        "icone": "🔧",
        "equipe": "Equipe Manutenção",
        "turno": "05:00 - 14:00",
    },
    {
        "valor": "Jefferson (Eletricista)",
        "nome": "Jefferson",
        "funcao": "Eletricista",
        "icone": "⚡",
        "equipe": "Equipe Manutenção",
        "turno": "05:00 - 14:00",
    },
    {
        "valor": "Edivaldo",
        "nome": "Edivaldo",
        "funcao": "Eletricista",
        "icone": "⚡",
        "equipe": "Equipe Manutenção",
        "turno": "07:00 - 16:00",
    },
    {
        "valor": "Erivaldo",
        "nome": "Erivaldo",
        "funcao": "Mecânico",
        "icone": "🔧",
        "equipe": "Equipe Manutenção",
        "turno": "07:00 - 16:00",
    },
    {
        "valor": "Carlos",
        "nome": "Carlos",
        "funcao": "Predial",
        "icone": "🏢",
        "equipe": "Manutenção Predial",
        "turno": "07:00 - 16:00",
    },
    {
        "valor": "Tiago",
        "nome": "Tiago",
        "funcao": "Predial",
        "icone": "🏢",
        "equipe": "Manutenção Predial",
        "turno": "07:00 - 16:00",
    },
    {
        "valor": "Ailton",
        "nome": "Ailton",
        "funcao": "Mecânico",
        "icone": "🔧",
        "equipe": "Equipe Manutenção",
        "turno": "15:30 - 01:04",
    },
    {
        "valor": "Douglas",
        "nome": "Douglas",
        "funcao": "Eletricista",
        "icone": "⚡",
        "equipe": "Equipe Manutenção",
        "turno": "15:30 - 01:04",
    },
    {
        "valor": "Jefferson (Aux. mecânico)",
        "nome": "Jefferson",
        "funcao": "Aux. mecânico",
        "icone": "🔧",
        "equipe": "Equipe Manutenção",
        "turno": "15:30 - 01:04",
    },
]

FUNCIONARIOS_VALIDOS = {f["valor"] for f in FUNCIONARIOS_EQUIPE} | {
    f["nome"] for f in FUNCIONARIOS_EQUIPE
}

FREQUENCIA_PCM = {
    "Semanal": ("dias", 7),
    "Quinzenal": ("dias", 15),
    "Mensal": ("meses", 1),
    "Bimestral": ("meses", 2),
    "Trimestral": ("meses", 3),
    "Semestral": ("meses", 6),
    "Anual": ("meses", 12),
}

MESES_PT = {
    "01": "Janeiro",
    "02": "Fevereiro",
    "03": "Março",
    "04": "Abril",
    "05": "Maio",
    "06": "Junho",
    "07": "Julho",
    "08": "Agosto",
    "09": "Setembro",
    "10": "Outubro",
    "11": "Novembro",
    "12": "Dezembro",
}

CENTROS_CUSTO_OFICIAIS = [
    "20111 – Coordenação",
    "20112 – Liderança",
    "20113 – Planejamento Programação e Controle da Produção",
    "20211 – Estoque",
    "20212 – Movimentação (Empilhadeira)",
    "20221 – Massa",
    "20231 – Cilindro",
    "20241 – Trafila (alma)",
    "20242 – Trançadeira",
    "20243 – Trafila (capa)",
    "20311 – Corte",
    "20312 – Rebarba",
    "20411 – Apoio operacional (ferramentas e mangueira cortada)",
    "20412 – Moldagem",
    "20511 – Calandra",
    "20521 – Montagem",
    "20522 – Corte",
    "20523 – Desmoldagem",
    "20531 – Montagem",
    "20611 – VULCANIZAÇÃO",
    "20711 – Separação",
    "20712 – Montagem",
    "20721 – Injeção (conexões)",
    "20722 – Sobre injeção (capa)",
    "20731 – Prensa",
    "20811 – Separação",
    "20812 – Etiquetagem",
    "20813 – Carimbo",
    "20814 – Embalagem",
    "20821 – Montagem",
    "20911 – Estoque",
    "20921 – Separação",
    "20922 – Conferência",
    "20923 – Expedição",
    "20931 – Transporte",
    "21011 – Projetos",
    "21012 – Laboratório",
    "21111 – Controle de Qualidade",
    "21112 – Assistência Técnica",
    "21211 – Predial",
    "21212 – Elétrica",
    "21213 – Mecânica",
    "21214 – Ferramentaria",
    "21221 – Caldeira",
    "21222 – Tratamento de Efluentes Industriais",
    "21311 – Portaria",
    "21312 – Limpeza",
    "21313 – Segurança Patrimonial",
    "21411 – Segurança do Trabalho",
    "21412 – Medicina do Trabalho",
    "215 – ÁREA ADMINISTRATIVA",
    "21525 – Administrativo",
    "21526 – Compras",
    "21527 – Faturamento",
    "21531 – Gerência",
    "21532 – Financeiro",
    "21541 – Gerência",
    "21542 – Vendas Representantes",
    "21543 – Administração de Vendas",
    "21551 – Gerência",
    "21552 – Marketing",
    "21553 – Trade marketing",
    "21561 – Gerência",
    "21562 – Sistema Operacional",
    "21563 – Tecnologia da Informação",
    "21571 – Departamento Pessoal",
]

FAVICON_J = (
    "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0"
    " 0 512 512'%3E%3Crect width='512' height='512' rx='96'"
    " fill='%23f97316'/%3E%3Cpath d='M153 143h200v190H323V203H153zm0 190c0 52 38"
    " 95 90 95s90-43 90-95h-50c0 24-18 45-40 45s-40-21-40-45z'"
    " fill='%23fff'/%3E%3C/svg%3E"
)

PWA_HEAD = f"""
<link rel="manifest" href="/manifest.json">
<meta name="theme-color" content="#f97316">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Manutenção JAMAICA">
<link rel="apple-touch-icon" href="{FAVICON_J}">
"""


# ==============================================================================
# MODELOS DE BANCO DE DADOS
# ==============================================================================
class OrdemServico(db.Model):
    __tablename__ = "ordens_servico"
    id = db.Column(db.Integer, primary_key=True)
    numero_os = db.Column(db.String(20), unique=True, nullable=False, index=True)
    solicitante = db.Column(db.String(100), nullable=False)
    setor = db.Column(db.String(100), default="")
    titulo = db.Column(db.String(200), default="")
    descricao = db.Column(db.Text, default="")
    prioridade = db.Column(db.String(30), default="Média")
    responsavel = db.Column(db.String(100), default="")
    abertura = db.Column(db.String(50), nullable=False)
    finalizacao = db.Column(db.String(50), default="")
    status = db.Column(db.String(50), default="Aberta", index=True)
    tempo_total = db.Column(db.String(50), default="")
    tipo_manutencao = db.Column(db.String(30), default="")
    funcionarios = db.Column(db.String(200), default="")
    horas_trabalhadas = db.Column(db.Float, default=0.0)
    observacao_tecnica = db.Column(db.Text, default="")
    motivo_pausa = db.Column(db.String(200), default="")
    evidencia_imagem = db.Column(db.Text, default="")
    evidencia_nome = db.Column(db.Text, default="")

    @property
    def evidencias_lista(self):
        return listar_evidencias_ordem(self)


class EvidenciaOS(db.Model):
    __tablename__ = "evidencias_os"
    id = db.Column(db.Integer, primary_key=True)
    ordem_id = db.Column(
        db.Integer,
        db.ForeignKey("ordens_servico.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )
    nome_arquivo = db.Column(db.String(220), unique=True, nullable=False, index=True)
    nome_original = db.Column(db.String(220), default="")
    mime_type = db.Column(db.String(80), default="image/jpeg")
    dados_binarios = db.Column(db.LargeBinary, nullable=False)
    criado_em = db.Column(db.String(50), default="")


class ItemEstoque(db.Model):
    __tablename__ = "itens_estoque"
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(50), unique=True, nullable=False, index=True)
    nome = db.Column(db.String(150), nullable=False)
    categoria = db.Column(db.String(100), default="Geral")
    localizacao = db.Column(db.String(100), default="")
    quantidade = db.Column(db.Integer, nullable=False, default=0)
    estoque_minimo = db.Column(db.Integer, nullable=False, default=2)
    ultima_atualizacao = db.Column(db.String(50), nullable=False)


class HistoricoMovimentacao(db.Model):
    __tablename__ = "historico_movimentacoes"
    id = db.Column(db.Integer, primary_key=True)
    peca_codigo = db.Column(db.String(50), nullable=False, index=True)
    peca_nome = db.Column(db.String(150), nullable=False)
    tipo = db.Column(db.String(20), nullable=False)
    quantidade = db.Column(db.Integer, nullable=False)
    solicitante = db.Column(db.String(100), default="")
    data_hora = db.Column(db.String(50), nullable=False)


class MaquinaSpare(db.Model):
    __tablename__ = "maquinas_spare"
    id = db.Column(db.Integer, primary_key=True)
    maquina = db.Column(db.String(150), nullable=False)
    modelo = db.Column(db.String(100), nullable=False)
    setor = db.Column(db.String(100), default="")
    pecas_necessarias = db.Column(db.Text, default="")
    pecas = db.relationship(
        "PecaSpare", backref="maquina_spare", cascade="all, delete-orphan"
    )


class PecaSpare(db.Model):
    __tablename__ = "pecas_spare"
    id = db.Column(db.Integer, primary_key=True)
    maquina_spare_id = db.Column(
        db.Integer, db.ForeignKey("maquinas_spare.id"), nullable=False
    )
    nome_peca = db.Column(db.String(150), nullable=False)
    codigo_peca = db.Column(db.String(50), default="")
    quantidade = db.Column(db.Integer, default=1)
    observacao = db.Column(db.String(200), default="")


class PlanoPreventivo(db.Model):
    __tablename__ = "planos_preventivos"
    id = db.Column(db.Integer, primary_key=True)
    maquina = db.Column(db.String(150), nullable=False)
    codigo_maquina = db.Column(db.String(50), default="")
    setor = db.Column(db.String(100), default="")
    atividade = db.Column(db.Text, nullable=False)
    frequencia = db.Column(db.String(50), default="Mensal")
    responsavel = db.Column(db.String(100), default="")
    proxima_execucao = db.Column(db.String(50), nullable=False)
    status = db.Column(db.String(30), default="Agendado")


class ArquivoPCM(db.Model):
    __tablename__ = "arquivos_pcm"
    id = db.Column(db.Integer, primary_key=True)
    nome_original = db.Column(db.String(220), nullable=False)
    nome_arquivo = db.Column(db.String(220), unique=True, nullable=False)
    tamanho_kb = db.Column(db.Float, default=0.0)
    enviado_por = db.Column(db.String(100), default="PCM")
    data_envio = db.Column(db.String(50), nullable=False)


class SolicitacaoCompra(db.Model):
    __tablename__ = "solicitacoes_compras"
    id = db.Column(db.Integer, primary_key=True)
    numero_sc = db.Column(db.String(100), nullable=False, index=True)
    tipo_sc = db.Column(db.String(50), default="S.C. Emergencial")
    solicitante = db.Column(db.String(100), nullable=False)
    setor_solicitante = db.Column(db.String(100), default="")
    responsaveis_duvidas = db.Column(db.String(150), default="")
    centro_custo = db.Column(db.String(150), default="")
    objeto_solicitacao = db.Column(db.String(250), nullable=False)
    data_solicitacao = db.Column(db.String(50), nullable=False)
    status_aprovacao = db.Column(db.String(50), default="Aguardando Aprovação")
    itens = db.relationship(
        "ItemSolicitacaoCompra",
        backref="solicitacao_compra",
        cascade="all, delete-orphan",
    )


class ItemSolicitacaoCompra(db.Model):
    __tablename__ = "itens_solicitacoes_compras"
    id = db.Column(db.Integer, primary_key=True)
    solicitacao_id = db.Column(
        db.Integer, db.ForeignKey("solicitacoes_compras.id"), nullable=False
    )
    descritivo = db.Column(db.Text, nullable=False)
    quantidade = db.Column(db.Integer, default=1)
    unidade = db.Column(db.String(20), default="Un")
    link_referencia = db.Column(db.Text, default="")


class LogAuditoria(db.Model):
    __tablename__ = "logs_auditoria"
    id = db.Column(db.Integer, primary_key=True)
    data_hora = db.Column(db.String(50), nullable=False)
    modulo = db.Column(db.String(50), nullable=False)
    acao = db.Column(db.String(100), nullable=False)
    detalhes = db.Column(db.Text, default="")


def registrar_auditoria(modulo, acao, detalhes=""):
    try:
        db.session.add(
            LogAuditoria(
                data_hora=agora_formatado(),
                modulo=modulo,
                acao=acao,
                detalhes=detalhes,
            )
        )
        db.session.commit()
    except Exception:
        db.session.rollback()


def backup_automatico_diario():
    if IS_POSTGRES or not os.path.exists(DB_PATH):
        return
    hoje_str = date.today().strftime("%Y%m%d")
    destino = os.path.join(BACKUP_DIR, f"os_backup_{hoje_str}.db")
    if not os.path.exists(destino):
        try:
            shutil.copy2(DB_PATH, destino)
            arquivos = sorted(
                [
                    os.path.join(BACKUP_DIR, f)
                    for f in os.listdir(BACKUP_DIR)
                    if f.endswith(".db")
                ],
                key=os.path.getmtime,
            )
            for antigo in arquivos[:-15]:
                os.remove(antigo)
        except OSError:
            pass


def sincronizar_evidencias_disco_com_banco():
    """Garante que fotos já existentes no disco sejam copiadas para o banco."""
    try:
        os.makedirs(EVIDENCIAS_DIR, exist_ok=True)
        for nome_arq in os.listdir(EVIDENCIAS_DIR):
            nome_limpo = os.path.basename(nome_arq)
            ext = os.path.splitext(nome_limpo)[1].lower()
            if ext not in EXTENSOES_EVIDENCIA:
                continue
            caminho = os.path.join(EVIDENCIAS_DIR, nome_limpo)
            if not os.path.isfile(caminho):
                continue
            if not EvidenciaOS.query.filter_by(nome_arquivo=nome_limpo).first():
                with open(caminho, "rb") as f:
                    dados = f.read()
                if dados:
                    ordem_match = OrdemServico.query.filter(
                        OrdemServico.evidencia_imagem.contains(nome_limpo)
                    ).first()
                    db.session.add(
                        EvidenciaOS(
                            ordem_id=ordem_match.id if ordem_match else None,
                            nome_arquivo=nome_limpo,
                            nome_original=nome_limpo,
                            mime_type=EXTENSOES_EVIDENCIA[ext],
                            dados_binarios=dados,
                            criado_em=agora_formatado(),
                        )
                    )
        db.session.commit()
    except Exception:
        db.session.rollback()


def inicializar_banco():
    db.create_all()
    try:
        insp = inspect(db.engine)
        cols = [c["name"] for c in insp.get_columns("ordens_servico")]
        if "observacao_tecnica" not in cols:
            db.session.execute(
                text(
                    "ALTER TABLE ordens_servico ADD COLUMN"
                    " observacao_tecnica TEXT DEFAULT ''"
                )
            )
        if "motivo_pausa" not in cols:
            db.session.execute(
                text(
                    "ALTER TABLE ordens_servico ADD COLUMN motivo_pausa"
                    " VARCHAR(200) DEFAULT ''"
                )
            )
        if "evidencia_imagem" not in cols:
            db.session.execute(
                text(
                    "ALTER TABLE ordens_servico ADD COLUMN evidencia_imagem"
                    " TEXT DEFAULT ''"
                )
            )
        elif IS_POSTGRES:
            db.session.execute(
                text(
                    "ALTER TABLE ordens_servico ALTER COLUMN evidencia_imagem"
                    " TYPE TEXT"
                )
            )
        if "evidencia_nome" not in cols:
            db.session.execute(
                text(
                    "ALTER TABLE ordens_servico ADD COLUMN evidencia_nome"
                    " TEXT DEFAULT ''"
                )
            )
        elif IS_POSTGRES:
            db.session.execute(
                text(
                    "ALTER TABLE ordens_servico ALTER COLUMN evidencia_nome"
                    " TYPE TEXT"
                )
            )
        db.session.commit()
    except Exception:
        db.session.rollback()
    sincronizar_evidencias_disco_com_banco()
    backup_automatico_diario()


with app.app_context():
    inicializar_banco()


# ==============================================================================
# ROTAS PWA (PROGRESSIVE WEB APP)
# ==============================================================================
@app.route("/manifest.json")
def manifest_pwa():
    return jsonify({
        "name": "Central de Manutenção JAMAICA",
        "short_name": "Manutenção",
        "description": (
            "Gestão de Ordens de Serviço, Estoque, PCM e Compras — JAMAICA"
        ),
        "start_url": "/os",
        "display": "standalone",
        "orientation": "any",
        "background_color": "#08090c",
        "theme_color": "#f97316",
        "icons": [{
            "src": FAVICON_J,
            "sizes": "192x192 512x512",
            "type": "image/svg+xml",
            "purpose": "any maskable",
        }],
    })


@app.route("/sw.js")
def service_worker():
    sw_code = """
    const CACHE_NAME = 'manutencao-jamaica-v6';
    self.addEventListener('install', event => {
        self.skipWaiting();
    });
    self.addEventListener('activate', event => {
        event.waitUntil(clients.claim());
    });
    self.addEventListener('fetch', event => {
        if (event.request.method !== 'GET') return;
        event.respondWith(
            fetch(event.request).catch(() => caches.match(event.request))
        );
    });
    """
    return Response(sw_code, mimetype="application/javascript")


# ==============================================================================
# FUNÇÕES UTILITÁRIAS, MÚLTIPLAS EVIDÊNCIAS E SEGURANÇA
# ==============================================================================
def agora_formatado():
    return datetime.now().strftime(FORMATO_DATA)


def hoje_iso():
    return date.today().strftime("%Y-%m-%d")


def _decodificar_campo_evidencia(raw):
    texto = (raw or "").strip()
    if not texto:
        return []
    if texto.startswith("["):
        try:
            lista = json.loads(texto)
            if isinstance(lista, list):
                return [str(x).strip() for x in lista if str(x).strip()]
        except Exception:
            pass
    return [texto]


def listar_evidencias_ordem(ordem):
    arquivos = _decodificar_campo_evidencia(
        getattr(ordem, "evidencia_imagem", "")
    )
    nomes = _decodificar_campo_evidencia(getattr(ordem, "evidencia_nome", ""))
    itens = []
    for idx, arq in enumerate(arquivos):
        arq_limpo = os.path.basename(arq)
        if not arq_limpo:
            continue
        nome_exib = nomes[idx] if idx < len(nomes) and nomes[idx] else arq_limpo
        itens.append({
            "arquivo": arq_limpo,
            "nome": nome_exib,
            "url": f"/os/evidencia/{arq_limpo}",
        })
    return itens


def gravar_evidencias_ordem(ordem, lista_itens):
    if not lista_itens:
        ordem.evidencia_imagem = ""
        ordem.evidencia_nome = ""
        return
    ordem.evidencia_imagem = json.dumps(
        [item["arquivo"] for item in lista_itens], ensure_ascii=False
    )
    ordem.evidencia_nome = json.dumps(
        [item["nome"] for item in lista_itens], ensure_ascii=False
    )


def salvar_evidencia_os(arquivo, numero_os, ordem_id=None):
    if not arquivo or not arquivo.filename:
        return "", "", None

    nome_orig = os.path.basename(arquivo.filename).strip()
    ext = os.path.splitext(nome_orig)[1].lower()
    if ext not in EXTENSOES_EVIDENCIA:
        return "", "", "Formato de imagem inválido. Use PNG, JPG, JPEG ou WEBP."

    conteudo = arquivo.read()
    if not conteudo:
        return "", "", "Um dos arquivos de evidência está vazio."
    if len(conteudo) > TAMANHO_MAX_EVIDENCIA:
        return (
            "",
            "",
            f"A imagem '{nome_orig[:40]}' excede o limite de 6 MB por foto.",
        )

    eh_png = conteudo.startswith(b"\x89PNG\r\n\x1a\n")
    eh_jpg = conteudo.startswith(b"\xff\xd8\xff")
    eh_webp = conteudo.startswith(b"RIFF") and conteudo[8:12] == b"WEBP"
    if not (eh_png or eh_jpg or eh_webp):
        return (
            "",
            "",
            f"O arquivo '{nome_orig[:40]}' não parece ser uma imagem válida.",
        )

    os.makedirs(EVIDENCIAS_DIR, exist_ok=True)
    slug_os = re.sub(r"[^A-Za-z0-9_-]", "_", numero_os or "os")
    nome_seguro = f"os_{slug_os}_{secrets.token_hex(6)}{ext}"
    destino = os.path.join(EVIDENCIAS_DIR, nome_seguro)
    with open(destino, "wb") as f:
        f.write(conteudo)

    db.session.add(
        EvidenciaOS(
            ordem_id=ordem_id,
            nome_arquivo=nome_seguro,
            nome_original=nome_orig[:200],
            mime_type=EXTENSOES_EVIDENCIA[ext],
            dados_binarios=conteudo,
            criado_em=agora_formatado(),
        )
    )

    return nome_seguro, nome_orig[:200], None


def salvar_multiplas_evidencias_os(arquivos, numero_os, existentes=0, ordem_id=None):
    validos = [a for a in (arquivos or []) if a and (a.filename or "").strip()]
    if not validos:
        return [], None

    if existentes + len(validos) > MAX_FOTOS_EVIDENCIA_POR_OS:
        return (
            [],
            f"Limite máximo de {MAX_FOTOS_EVIDENCIA_POR_OS} fotos por O.S. atingido.",
        )

    salvos = []
    for arq in validos:
        nome_seguro, nome_orig, erro = salvar_evidencia_os(
            arq, numero_os, ordem_id=ordem_id
        )
        if erro:
            for item in salvos:
                remover_evidencia_os(item["arquivo"])
            return [], erro
        salvos.append({
            "arquivo": nome_seguro,
            "nome": nome_orig,
            "url": f"/os/evidencia/{nome_seguro}",
        })
    return salvos, None


def remover_evidencia_os(nome_arquivo):
    if not nome_arquivo:
        return
    for item in _decodificar_campo_evidencia(nome_arquivo):
        nome_limpo = os.path.basename(item)
        try:
            EvidenciaOS.query.filter_by(nome_arquivo=nome_limpo).delete()
            db.session.commit()
        except Exception:
            db.session.rollback()
        caminho = os.path.join(EVIDENCIAS_DIR, nome_limpo)
        if os.path.exists(caminho):
            try:
                os.remove(caminho)
            except OSError:
                pass


def gerar_numero_os():
    numeros = OrdemServico.query.with_entities(OrdemServico.numero_os).all()
    maior = NUMERO_OS_INICIAL - 1
    for (numero,) in numeros:
        try:
            maior = max(maior, int(numero.split("-")[0]))
        except (ValueError, AttributeError, IndexError):
            pass
    return f"{maior + 1}-{PREFIXO_ANO_OS}"


def gerar_numero_sc():
    numeros = SolicitacaoCompra.query.with_entities(
        SolicitacaoCompra.numero_sc
    ).all()
    maior = NUMERO_SC_INICIAL - 1
    for (numero,) in numeros:
        if not numero:
            continue
        match = re.search(r"(\d{4})", numero)
        if match:
            try:
                maior = max(maior, int(match.group(1)))
            except ValueError:
                pass
    proximo = maior + 1
    return f"01 – S.C. – {proximo:04d} – {PREFIXO_ANO_SC}"


def calcular_tempo_total(abertura, finalizacao):
    inicio = datetime.strptime(abertura, FORMATO_DATA)
    fim = datetime.strptime(finalizacao, FORMATO_DATA)
    segundos = max(0, int((fim - inicio).total_seconds()))
    dias, resto = divmod(segundos, 86400)
    horas, resto = divmod(resto, 3600)
    minutos, _ = divmod(resto, 60)
    partes = []
    if dias:
        partes.append(f"{dias}d")
    if horas:
        partes.append(f"{horas}h")
    partes.append(f"{minutos}min")
    return " ".join(partes)


def tempo_hhmm_para_horas_decimais(hhmm_str):
    if not hhmm_str or ":" not in hhmm_str:
        return 0.0
    try:
        partes = hhmm_str.split(":")
        h = int(partes[0])
        m = int(partes[1])
        return round(h + (m / 60.0), 2)
    except (ValueError, IndexError):
        return 0.0


def horas_decimais_para_hhmm(horas_float):
    if not horas_float or horas_float <= 0:
        return "0h 00min"
    h = int(horas_float)
    m = int(round((horas_float - h) * 60))
    if m == 60:
        h += 1
        m = 0
    return f"{h}h {m:02d}min"


def horas_decimais_para_input_time(horas_float):
    if not horas_float or horas_float <= 0:
        return "01:00"
    total_min = int(round(horas_float * 60))
    h = min(99, total_min // 60)
    m = total_min % 60
    return f"{h:02d}:{m:02d}"


def formatar_duracao_indicador(horas_float):
    if not horas_float or horas_float <= 0:
        return "0h 00min"
    if horas_float >= 24:
        return f"{round(horas_float / 24, 1)} dias"
    return horas_decimais_para_hhmm(horas_float)


def normalizar_chave_maquina(nome):
    return re.sub(r"\s+", " ", str(nome or "").strip()).lower()


def token_csrf():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_hex(16)
    return session["csrf_token"]


@app.before_request
def proteger_csrf():
    if request.method != "POST":
        return None
    enviado = request.form.get("csrf_token", "") or request.headers.get(
        "X-CSRF-Token", ""
    )
    esperado = session.get("csrf_token", "")
    if not esperado or not hmac.compare_digest(
        enviado.encode(), esperado.encode()
    ):
        if (
            request.is_json
            or request.headers.get("X-Requested-With") == "XMLHttpRequest"
        ):
            return (
                jsonify(
                    {"ok": False, "erro": "Token CSRF inválido. Recarregue a página."}
                ),
                400,
            )
        flash(
            "Sessão expirada ou requisição inválida. Recarregue a página e tente"
            " novamente."
        )
        return redirect(request.referrer or url_for("dashboard"))
    return None


@app.errorhandler(413)
def arquivo_muito_grande(e):
    flash("Os arquivos enviados excedem o tamanho máximo permitido (48 MB).")
    return redirect(request.referrer or url_for("dashboard"))


@app.after_request
def adicionar_headers_seguranca(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    return response


def parse_data_br(texto):
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return None
    if isinstance(texto, datetime):
        return texto.date()
    if isinstance(texto, date):
        return texto
    s = str(texto).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], fmt[: len(s[:19])]).date()
        except ValueError:
            continue
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass
    return None


def somar_meses(data, meses):
    indice = data.month - 1 + meses
    ano = data.year + indice // 12
    mes = indice % 12 + 1
    dia = min(data.day, calendar.monthrange(ano, mes)[1])
    return date(ano, mes, dia)


def avancar_data(data, frequencia):
    tipo, valor = FREQUENCIA_PCM.get(frequencia, ("meses", 1))
    if tipo == "dias":
        return data + timedelta(days=valor)
    return somar_meses(data, valor)


def status_plano(plano, hoje=None):
    hoje = hoje or date.today()
    data = parse_data_br(plano.proxima_execucao)
    if data is not None and data < hoje:
        return "Atrasado"
    return "Agendado"


def chave_mes_os(ordem):
    for campo in (ordem.finalizacao, ordem.abertura):
        try:
            return datetime.strptime(campo, FORMATO_DATA).strftime("%Y-%m")
        except (ValueError, TypeError):
            continue
    return None


def renderizar(template, **contexto):
    js = "window.__CSRF__=" + json.dumps(token_csrf()) + ";" + JS_BASE
    return render_template_string(
        template,
        css=CSS_BASE,
        js=js,
        favicon=FAVICON_J,
        pwa_head=PWA_HEAD,
        flash_msgs=get_flashed_messages(),
        **contexto,
    )


def _estilizar_aba_excel(ws, df, coluna_status=None):
    ws.freeze_panes = "A2"
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill(
        start_color="1F2937", end_color="1F2937", fill_type="solid"
    )
    accent_border = Side(border_style="medium", color="F97316")
    thin_border_side = Side(border_style="thin", color="D1D5DB")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="Calibri", size=10, color="111827")
    zebra_fill = PatternFill(
        start_color="F9FAFB", end_color="F9FAFB", fill_type="solid"
    )
    alert_fill = PatternFill(
        start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"
    )
    alert_font = Font(name="Calibri", size=10, bold=True, color="991B1B")
    ok_fill = PatternFill(
        start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"
    )
    ok_font = Font(name="Calibri", size=10, bold=True, color="166534")

    ws.row_dimensions[1].height = 26
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=True
        )
        cell.border = Border(
            bottom=accent_border, left=thin_border_side, right=thin_border_side
        )

    status_col_idx = None
    if coluna_status and coluna_status in df.columns:
        status_col_idx = list(df.columns).index(coluna_status) + 1

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        ws.row_dimensions[row_idx].height = 20
        is_even = row_idx % 2 == 0
        for col_idx, cell in enumerate(row, start=1):
            cell.font = cell_font
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            cell.border = Border(
                bottom=thin_border_side,
                left=thin_border_side,
                right=thin_border_side,
            )
            if is_even:
                cell.fill = zebra_fill
            if status_col_idx and col_idx == status_col_idx:
                val_str = str(cell.value or "").lower()
                if any(
                    k in val_str
                    for k in (
                        "baixo",
                        "atrasado",
                        "emergencial",
                        "reprovado",
                        "crítica",
                        "aguardando",
                    )
                ):
                    cell.fill = alert_fill
                    cell.font = alert_font
                elif any(
                    k in val_str
                    for k in ("normal", "finalizada", "aprovado", "comprado")
                ):
                    cell.fill = ok_fill
                    cell.font = ok_font

    for coluna in ws.columns:
        letra = coluna[0].column_letter
        maior = max([len(str(c.value or "")) for c in coluna] or [0])
        ws.column_dimensions[letra].width = min(max(maior + 4, 14), 45)


def gerar_excel(df, nome_planilha, prefixo, coluna_status=None):
    memoria = BytesIO()
    with pd.ExcelWriter(memoria, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nome_planilha)
        ws = writer.sheets[nome_planilha]
        _estilizar_aba_excel(ws, df, coluna_status=coluna_status)

    memoria.seek(0)
    nome = f"{prefixo}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    return send_file(
        memoria,
        as_attachment=True,
        download_name=nome,
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )


def gerar_excel_multiplas_abas(abas_config, prefixo):
    memoria = BytesIO()
    with pd.ExcelWriter(memoria, engine="openpyxl") as writer:
        for nome_aba, df, col_status in abas_config:
            titulo_limpo = nome_aba[:31]
            df.to_excel(writer, index=False, sheet_name=titulo_limpo)
            ws = writer.sheets[titulo_limpo]
            _estilizar_aba_excel(ws, df, coluna_status=col_status)

    memoria.seek(0)
    nome = f"{prefixo}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    return send_file(
        memoria,
        as_attachment=True,
        download_name=nome,
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )


def serializar_banco_completo():
    ordens = OrdemServico.query.order_by(OrdemServico.id.asc()).all()
    estoque = ItemEstoque.query.order_by(ItemEstoque.id.asc()).all()
    movs = HistoricoMovimentacao.query.order_by(
        HistoricoMovimentacao.id.asc()
    ).all()
    maquinas = MaquinaSpare.query.order_by(MaquinaSpare.id.asc()).all()
    planos = PlanoPreventivo.query.order_by(PlanoPreventivo.id.asc()).all()
    arquivos_pcm = ArquivoPCM.query.order_by(ArquivoPCM.id.asc()).all()
    compras = SolicitacaoCompra.query.order_by(SolicitacaoCompra.id.asc()).all()
    logs = LogAuditoria.query.order_by(LogAuditoria.id.asc()).all()

    return {
        "sistema": "Central de Manutenção JAMAICA",
        "exportado_em": agora_formatado(),
        "motor_banco": "PostgreSQL" if IS_POSTGRES else "SQLite",
        "resumo": {
            "ordens_servico": len(ordens),
            "itens_estoque": len(estoque),
            "movimentacoes": len(movs),
            "maquinas_spare": len(maquinas),
            "planos_pcm": len(planos),
            "arquivos_pcm": len(arquivos_pcm),
            "solicitacoes_compras": len(compras),
            "logs_auditoria": len(logs),
        },
        "ordens_servico": [
            {
                "id": o.id,
                "numero_os": o.numero_os,
                "solicitante": o.solicitante,
                "setor": o.setor,
                "titulo": o.titulo,
                "descricao": o.descricao,
                "prioridade": o.prioridade,
                "responsavel": o.responsavel,
                "abertura": o.abertura,
                "finalizacao": o.finalizacao,
                "status": o.status,
                "tempo_total": o.tempo_total,
                "tipo_manutencao": o.tipo_manutencao,
                "funcionarios": o.funcionarios,
                "horas_trabalhadas": o.horas_trabalhadas,
                "observacao_tecnica": o.observacao_tecnica,
                "motivo_pausa": o.motivo_pausa,
                "evidencias": listar_evidencias_ordem(o),
            }
            for o in ordens
        ],
        "itens_estoque": [
            {
                "id": p.id,
                "codigo": p.codigo,
                "nome": p.nome,
                "categoria": p.categoria,
                "localizacao": p.localizacao,
                "quantidade": p.quantidade,
                "estoque_minimo": p.estoque_minimo,
                "ultima_atualizacao": p.ultima_atualizacao,
            }
            for p in estoque
        ],
        "historico_movimentacoes": [
            {
                "id": m.id,
                "peca_codigo": m.peca_codigo,
                "peca_nome": m.peca_nome,
                "tipo": m.tipo,
                "quantidade": m.quantidade,
                "solicitante": m.solicitante,
                "data_hora": m.data_hora,
            }
            for m in movs
        ],
        "maquinas_spare": [
            {
                "id": m.id,
                "maquina": m.maquina,
                "modelo": m.modelo,
                "setor": m.setor,
                "pecas_necessarias": m.pecas_necessarias,
                "pecas": [
                    {
                        "id": ps.id,
                        "nome_peca": ps.nome_peca,
                        "codigo_peca": ps.codigo_peca,
                        "quantidade": ps.quantidade,
                        "observacao": ps.observacao,
                    }
                    for ps in m.pecas
                ],
            }
            for m in maquinas
        ],
        "planos_preventivos": [
            {
                "id": pl.id,
                "maquina": pl.maquina,
                "codigo_maquina": pl.codigo_maquina,
                "setor": pl.setor,
                "atividade": pl.atividade,
                "frequencia": pl.frequencia,
                "responsavel": pl.responsavel,
                "proxima_execucao": pl.proxima_execucao,
                "status": pl.status,
            }
            for pl in planos
        ],
        "arquivos_pcm": [
            {
                "id": aq.id,
                "nome_original": aq.nome_original,
                "nome_arquivo": aq.nome_arquivo,
                "tamanho_kb": aq.tamanho_kb,
                "enviado_por": aq.enviado_por,
                "data_envio": aq.data_envio,
            }
            for aq in arquivos_pcm
        ],
        "solicitacoes_compras": [
            {
                "id": sc.id,
                "numero_sc": sc.numero_sc,
                "tipo_sc": sc.tipo_sc,
                "solicitante": sc.solicitante,
                "setor_solicitante": sc.setor_solicitante,
                "responsaveis_duvidas": sc.responsaveis_duvidas,
                "centro_custo": sc.centro_custo,
                "objeto_solicitacao": sc.objeto_solicitacao,
                "data_solicitacao": sc.data_solicitacao,
                "status_aprovacao": sc.status_aprovacao,
                "itens": [
                    {
                        "id": it.id,
                        "descritivo": it.descritivo,
                        "quantidade": it.quantidade,
                        "unidade": it.unidade,
                        "link_referencia": it.link_referencia,
                    }
                    for it in sc.itens
                ],
            }
            for sc in compras
        ],
        "logs_auditoria": [
            {
                "id": lg.id,
                "data_hora": lg.data_hora,
                "modulo": lg.modulo,
                "acao": lg.acao,
                "detalhes": lg.detalhes,
            }
            for lg in logs
        ],
    }


# ==============================================================================
# ESTILOS E SCRIPTS BASE (LIMPOS E RESPONSIVOS EM DESKTOP E MOBILE)
# ==============================================================================
CSS_BASE = r"""
:root {
    --bg: #08090c;
    --s1: #111318;
    --s2: #181b22;
    --s3: #20242e;
    --bd: #2a2f3d;
    --tx: #f3f4f6;
    --mt: #9ca3af;
    --or: #f97316;
    --or2: #fb923c;
    --er: #ef4444;
    --ok: #22c55e;
    --info: #3b82f6;
    --grad: linear-gradient(135deg, #090a0d 0%, #431407 52%, #ea580c 100%);
    --btn: linear-gradient(135deg, #ea580c, #f97316 60%, #fb923c);
    --r: 14px;
}
* { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
html { scroll-behavior: smooth; overflow-x: clip; }
body {
    margin: 0;
    padding: 20px;
    min-height: 100vh;
    font-family: "Inter", "Segoe UI", system-ui, -apple-system, sans-serif;
    color: var(--tx);
    background:
        radial-gradient(950px 520px at 88% -8%, rgba(249, 115, 22, 0.20), transparent 62%),
        radial-gradient(650px 400px at 5% 100%, rgba(59, 130, 246, 0.05), transparent 60%),
        var(--bg);
    background-attachment: fixed;
    line-height: 1.45;
    overflow-x: clip;
}
.container { max-width: 1640px; width: 100%; margin: auto; animation: fadeIn .3s ease-out; overflow-x: clip; }
.muted { color: var(--mt); font-size: 13px; }

/* LOBBY DE ENTRADA */
.lobby-wrap { max-width: 820px; margin: 4vh auto; }
.lobby-card {
    background: linear-gradient(175deg, var(--s2), var(--s1));
    border: 1px solid rgba(251, 146, 60, 0.32);
    border-radius: 18px;
    padding: 32px;
    text-align: center;
    box-shadow: 0 22px 60px rgba(0, 0, 0, 0.65), 0 0 40px rgba(249, 115, 22, 0.14);
}
.lobby-logo {
    width: 68px; height: 68px; margin: 0 auto 16px; border-radius: 16px;
    background: var(--btn); display: flex; align-items: center; justify-content: center;
    font-size: 32px; font-weight: 900; color: #fff;
    box-shadow: 0 10px 28px rgba(249, 115, 22, 0.35);
}
.lobby-card h1 { margin: 0 0 8px; font-size: 26px; font-weight: 800; color: #fff; }
.lobby-card p.sub { margin: 0 auto 24px; max-width: 560px; color: #fed7aa; font-size: 14px; }
.lobby-acoes { display: flex; gap: 12px; justify-content: center; flex-wrap: wrap; margin-top: 24px; }
.btn-entrar-lobby {
    padding: 14px 28px !important; font-size: 15px !important; border-radius: 12px !important;
    background: var(--btn); color: #fff; box-shadow: 0 8px 24px rgba(249, 115, 22, 0.4);
}

/* CABEÇALHO E NAVEGAÇÃO */
.cabecalho {
    background: var(--grad);
    padding: 22px 26px;
    border-radius: var(--r);
    margin-bottom: 20px;
    border: 1px solid rgba(251, 146, 60, 0.28);
    box-shadow: 0 16px 45px rgba(0, 0, 0, 0.55), 0 0 30px rgba(249, 115, 22, 0.14);
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    flex-wrap: wrap;
    gap: 14px;
    position: relative;
    overflow: hidden;
}
.cabecalho-principal { flex: 1 1 auto; min-width: 0; }
.cabecalho-titulo { display: flex; align-items: center; gap: 14px; margin-bottom: 4px; min-width: 0; }
.cabecalho-texto { min-width: 0; flex: 1; }
.logo-badge {
    width: 42px; height: 42px; border-radius: 10px;
    background: rgba(0, 0, 0, 0.45); border: 1px solid rgba(255, 255, 255, 0.22);
    display: flex; align-items: center; justify-content: center;
    font-weight: 900; font-size: 20px; color: #fff; flex-shrink: 0;
}
.cabecalho h1 {
    margin: 0 0 3px; font-size: 21px; font-weight: 800; letter-spacing: 0.4px; color: #fff;
    overflow-wrap: break-word; line-height: 1.25;
}
.cabecalho p { margin: 0; color: #fed7aa; font-size: 13px; overflow-wrap: break-word; }
.cabecalho-lado { display: flex; flex-direction: column; align-items: flex-end; gap: 10px; flex-shrink: 0; }
.relogio {
    background: rgba(0, 0, 0, 0.52); padding: 7px 14px; border-radius: 99px;
    font-size: 12.5px; font-weight: 600; color: #fed7aa;
    border: 1px solid rgba(255, 255, 255, 0.14); font-variant-numeric: tabular-nums;
}
.nav-abas { display: flex; gap: 8px; margin-top: 14px; flex-wrap: wrap; width: 100%; }
.aba, .botao-sair {
    display: inline-flex; align-items: center; gap: 6px; padding: 9px 14px;
    background: rgba(0, 0, 0, 0.48); border: 1px solid rgba(255, 255, 255, 0.18);
    color: #f3f4f6; border-radius: 10px; text-decoration: none;
    font-size: 13px; font-weight: 600; white-space: nowrap;
    transition: all .2s ease; cursor: pointer; max-width: 100%;
}
.aba:hover { transform: translateY(-2px); border-color: var(--or2); background: rgba(0,0,0,0.7); }
.aba.ativa {
    background: #fff; color: #111827; border-color: #fff;
    font-weight: 800; box-shadow: 0 4px 16px rgba(0, 0, 0, 0.35);
}
.botao-sair { border-color: rgba(239, 68, 68, 0.55); color: #fecaca; background: rgba(127, 29, 29, 0.4); }
.botao-sair:hover { background: rgba(185, 28, 28, 0.65); transform: translateY(-2px); }

/* BARRA DE NAVEGAÇÃO INFERIOR MOBILE */
.mobile-bottom-nav { display: none; }

/* SUB-ABAS DA CENTRAL DE GESTÃO */
.sub-abas-nav {
    display: flex; gap: 8px; background: var(--s1); padding: 10px;
    border-radius: 12px; border: 1px solid var(--bd); margin-bottom: 20px; flex-wrap: wrap;
}
.sub-aba-btn {
    padding: 9px 16px; background: #0c0e12; border: 1px solid var(--bd);
    color: var(--mt); border-radius: 8px; font-size: 13px; font-weight: 700;
    cursor: pointer; transition: all .2s;
}
.sub-aba-btn:hover { border-color: var(--or); color: #fff; }
.sub-aba-btn.ativa {
    background: var(--btn); color: #fff; border-color: var(--or);
    box-shadow: 0 4px 14px rgba(249, 115, 22, 0.3);
}
.painel-secao { display: none; min-width: 0; }
.painel-secao.ativo { display: block; animation: fadeIn .25s ease; }

/* CARDS DE KPI */
.resumo {
    display: grid; grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 16px; margin-bottom: 20px;
}
.resumo-card, .card {
    background: linear-gradient(175deg, var(--s2), var(--s1));
    border: 1px solid var(--bd); border-radius: var(--r);
    padding: 20px 22px; margin-bottom: 20px;
    box-shadow: 0 8px 28px rgba(0, 0, 0, 0.32);
    transition: all .22s ease; min-width: 0;
}
.resumo-card { margin-bottom: 0; position: relative; cursor: default; }
.resumo-card[data-atalho] { cursor: pointer; }
.resumo-card:hover {
    transform: translateY(-2px); border-color: rgba(249, 115, 22, 0.6);
    box-shadow: 0 12px 32px rgba(249, 115, 22, 0.15);
}
.resumo-card span {
    display: block; color: var(--mt); font-size: 12px; font-weight: 600;
    margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.6px;
    padding-right: 28px;
}
.resumo-card strong { font-size: 28px; font-weight: 800; color: var(--or2); font-variant-numeric: tabular-nums; }
.resumo-card small { display: block; margin-top: 6px; color: var(--mt); font-size: 12px; }
.icone { position: absolute; right: 16px; top: 16px; font-size: 21px; opacity: .75; }
.progresso { height: 7px; background: #262a36; border-radius: 99px; margin-top: 10px; overflow: hidden; }
.progresso span { display: block; height: 100%; background: var(--btn); transition: width .8s ease; padding-right: 0; }

/* GRIDS */
.grid-duplo { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 20px; margin-bottom: 20px; }
.grid-triplo { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; }

/* SEÇÕES E FORMULÁRIOS */
.card h2 {
    margin: 0 0 14px; font-size: 17px; font-weight: 700;
    display: flex; gap: 10px; align-items: center;
}
.card h2:before {
    content: ""; width: 5px; height: 20px; border-radius: 4px; flex-shrink: 0;
    background: var(--btn); box-shadow: 0 0 10px rgba(249, 115, 22, 0.5);
}
details.card > summary {
    list-style: none; cursor: pointer; font-size: 16.5px; font-weight: 700;
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
}
details.card > summary::-webkit-details-marker { display: none; }
details.card > summary:after {
    content: "▾"; color: var(--or2); font-size: 18px; transition: transform .25s; flex-shrink: 0;
}
details.card[open] > summary:after { transform: rotate(180deg); }
details.card > .conteudo { margin-top: 16px; padding-top: 16px; border-top: 1px solid var(--bd); }

.form-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px 18px; }
.form-grid.tres { grid-template-columns: repeat(3, minmax(0, 1fr)); }
.campo { min-width: 0; }
.campo.completo { grid-column: 1 / -1; }
label { display: block; font-weight: 600; font-size: 12.5px; margin-bottom: 6px; color: #d1d5db; }
input, textarea, select {
    display: block; width: 100%; padding: 10px 13px;
    border: 1px solid var(--bd); border-radius: 10px;
    font: inherit; font-size: 13.5px; background: #0c0e12; color: var(--tx);
    transition: border-color .2s, box-shadow .2s;
}
input[type="file"] { padding: 8px 10px; background: #0c0e12; color: #d1d5db; cursor: pointer; }
input:focus, textarea:focus, select:focus {
    outline: none; border-color: var(--or);
    box-shadow: 0 0 0 3px rgba(249, 115, 22, 0.22);
}
input[disabled] { color: var(--or2); font-weight: 700; background: #14171f; opacity: 1; }
textarea { min-height: 86px; resize: vertical; }

.seg { display: flex; gap: 8px; flex-wrap: wrap; }
.seg input { display: none; }
.seg label {
    margin: 0; flex: 1; min-width: 85px; text-align: center; padding: 10px 12px;
    border: 1px solid var(--bd); border-radius: 10px; background: #0c0e12;
    color: var(--mt); cursor: pointer; font-weight: 600; transition: all .2s;
}
.seg input:checked + label {
    background: var(--btn); color: #fff; border-color: var(--or);
    box-shadow: 0 4px 14px rgba(249, 115, 22, 0.3);
}

/* BOTÕES */
button, .botao {
    display: inline-flex; align-items: center; justify-content: center; gap: 6px;
    padding: 10px 15px; border: 0; border-radius: 10px;
    font-size: 13px; font-weight: 700; font-family: inherit;
    text-decoration: none; cursor: pointer; color: #fff;
    transition: all .18s ease;
}
.salvar, .pesquisar, .finalizar, .adicionar, .movimentar {
    background: var(--btn); box-shadow: 0 4px 14px rgba(249, 115, 22, 0.25);
}
.excel {
    background: linear-gradient(135deg, #065f46, #10b981);
    box-shadow: 0 4px 14px rgba(16, 185, 129, 0.22);
}
.limpar { background: #232733; color: #e5e7eb; border: 1px solid var(--bd); }
.pausar { background: linear-gradient(135deg, #92400e, #d97706); }
.imprimir { background: linear-gradient(135deg, #1d4ed8, #3b82f6); }
.evidencia-btn { background: linear-gradient(135deg, #0f766e, #14b8a6); }
.excluir {
    background: linear-gradient(135deg, #7f1d1d, #dc2626);
    box-shadow: 0 4px 12px rgba(220, 38, 38, 0.22);
}
.editar-peca { background: #2c3140; border: 1px solid #4b5563; color: #f3f4f6; }
button:hover, .botao:hover { transform: translateY(-1px); filter: brightness(1.08); }
button:active, .botao:active { transform: translateY(0); }

/* PESQUISA E FILTROS */
.acoes-topo { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px; }
.busca { display: flex; gap: 8px; flex: 1; min-width: 240px; }
.campo-busca { position: relative; flex: 1; min-width: 0; }
.campo-busca input { padding-left: 38px; }
.campo-busca:before { content: "🔍"; position: absolute; left: 13px; top: 50%; transform: translateY(-50%); font-size: 13px; }
.filtros { display: flex; flex-wrap: wrap; gap: 9px; align-items: center; margin-top: 14px; }
.rotulo { font-size: 11px; font-weight: 700; text-transform: uppercase; color: var(--mt); letter-spacing: 0.6px; }
.chip {
    padding: 7px 13px; border-radius: 99px; background: #0c0e12;
    border: 1px solid var(--bd); color: var(--mt); font-size: 12.5px; cursor: pointer;
}
.chip:hover { border-color: var(--or); color: #fff; }
.chip.ativo { background: var(--btn); color: #fff; border-color: var(--or); }
.filtros select { width: auto; padding: 7px 13px; border-radius: 99px; font-size: 12.5px; }
.contagem { margin-left: auto; font-size: 13px; color: var(--mt); }
.contagem strong { color: var(--or2); }

/* TABELAS */
.tabela-wrapper {
    width: 100%; overflow-x: auto; border-radius: 12px;
    border: 1px solid var(--bd); background: var(--s1);
    -webkit-overflow-scrolling: touch;
}
table { width: 100%; min-width: 920px; border-collapse: collapse; }
th, td {
    padding: 11px 12px; border-bottom: 1px solid #1f232e;
    text-align: left; font-size: 13px; vertical-align: middle;
    overflow-wrap: break-word;
}
th {
    background: #0c0e12; color: var(--or2); white-space: nowrap;
    text-transform: uppercase; font-size: 11px; font-weight: 700; letter-spacing: .7px;
}
th.ordenavel { cursor: pointer; user-select: none; }
th.ordenavel:hover { color: #fff; }
tbody tr.linha { transition: background .15s; }
tbody tr.linha:hover { background: rgba(249, 115, 22, 0.07) !important; }
tbody tr.clicavel { cursor: pointer; }
tr.detalhe { display: none; }
tr.detalhe td { background: #0c0e12; padding: 14px 20px; border-left: 4px solid var(--or); }
tr.detalhe p { white-space: pre-wrap; margin: 6px 0 0; overflow-wrap: break-word; }
.vazio { text-align: center; padding: 28px !important; color: var(--mt); }
tr.sem-resultado { display: none; }

/* ARQUIVOS EXCEL PCM SALVOS */
.lista-arquivos-pcm {
    display: flex; flex-direction: column; gap: 10px; margin-top: 14px;
}
.item-arquivo-pcm {
    display: flex; align-items: center; justify-content: space-between; gap: 12px;
    flex-wrap: wrap; background: #0c0e12; border: 1px solid var(--bd);
    border-radius: 10px; padding: 12px 14px;
}
.item-arquivo-info { min-width: 0; flex: 1; }
.item-arquivo-info strong { display: block; color: var(--tx); font-size: 13.5px; overflow-wrap: break-word; }
.item-arquivo-info small { color: var(--mt); font-size: 12px; }
.item-arquivo-acoes { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }

/* GALERIA DE EVIDÊNCIAS FOTOGRÁFICAS */
.evidencia-box {
    margin-top: 12px; padding: 12px; background: #14171f;
    border: 1px solid var(--bd); border-radius: 10px;
}
.galeria-evidencias {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(170px, 1fr));
    gap: 10px; margin-top: 10px;
}
.evidencia-item-card {
    background: #0c0e12; border: 1px solid var(--bd); border-radius: 8px;
    padding: 8px; display: flex; flex-direction: column; gap: 6px;
}
.evidencia-item-card img {
    width: 100%; height: 130px; object-fit: cover; border-radius: 6px;
    border: 1px solid #262a36; display: block;
}
.evidencia-item-nome {
    font-size: 11px; color: var(--mt); white-space: nowrap;
    overflow: hidden; text-overflow: ellipsis;
}
.preview-evidencia-box {
    margin-top: 10px; padding: 10px; background: #0c0e12;
    border: 1px dashed var(--bd); border-radius: 10px; display: none;
}
.preview-grid {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(110px, 1fr));
    gap: 8px; margin-top: 8px;
}
.preview-grid-item {
    background: #14171f; border: 1px solid var(--bd);
    border-radius: 6px; padding: 4px; text-align: center;
}
.preview-grid-item img { width: 100%; height: 80px; object-fit: cover; border-radius: 4px; display: block; }
.preview-grid-item span {
    display: block; font-size: 10px; color: var(--mt); margin-top: 3px;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}

/* BADGES E STATUS */
.status {
    display: inline-block; padding: 5px 11px; border-radius: 99px;
    font-size: 11.5px; font-weight: 700; white-space: nowrap;
}
.aberta { background: rgba(250, 204, 21, 0.14); color: #fde047; border: 1px solid rgba(250, 204, 21, 0.32); }
.aguardando-peca { background: rgba(168, 85, 247, 0.18); color: #d8b4fe; border: 1px solid rgba(168, 85, 247, 0.38); }
.finalizada, .estoque-ok { background: rgba(34, 197, 94, 0.14); color: #4ade80; border: 1px solid rgba(34, 197, 94, 0.32); }
.estoque-baixo { background: rgba(239, 68, 68, 0.18); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.35); }
.vencido { color: #f87171; }

.prioridade { font-weight: 700; white-space: nowrap; }
.prioridade-critica { color: #f87171; }
.prioridade-alta { color: #fb923c; }
.prioridade-media { color: #facc15; }
.prioridade-baixa { color: #4ade80; }

.tipo-badge { display: inline-block; padding: 4px 10px; border-radius: 6px; font-size: 11.5px; font-weight: 700; white-space: nowrap; }
.tipo-preventiva { background: rgba(59, 130, 246, 0.15); color: #60a5fa; border: 1px solid rgba(59, 130, 246, 0.3); }
.tipo-corretiva { background: rgba(249, 115, 22, 0.15); color: #fb923c; border: 1px solid rgba(249, 115, 22, 0.3); }
.tipo-predial { background: rgba(16, 185, 129, 0.15); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); }

.sc-badge-emergencial, .sc-badge-programada, .sc-badge-reposicao {
    display: inline-block; padding: 5px 11px; border-radius: 99px;
    font-size: 11.5px; font-weight: 700; white-space: nowrap;
}
.sc-badge-emergencial { background: rgba(239, 68, 68, 0.18); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.38); }
.sc-badge-programada { background: rgba(59, 130, 246, 0.18); color: #60a5fa; border: 1px solid rgba(59, 130, 246, 0.38); }
.sc-badge-reposicao { background: rgba(168, 85, 247, 0.18); color: #c084fc; border: 1px solid rgba(168, 85, 247, 0.38); }

.select-status-sc {
    padding: 6px 10px; border-radius: 99px; font-size: 12px; font-weight: 700;
    cursor: pointer; border: 1px solid var(--bd); outline: none; width: auto;
}
.select-status-sc[data-val="Aprovado"] { background: rgba(34, 197, 94, 0.18); color: #4ade80; border-color: rgba(34, 197, 94, 0.4); }
.select-status-sc[data-val="Aguardando Aprovação"] { background: rgba(250, 204, 21, 0.18); color: #fde047; border-color: rgba(250, 204, 21, 0.4); }
.select-status-sc[data-val="Reprovado"] { background: rgba(239, 68, 68, 0.18); color: #f87171; border-color: rgba(239, 68, 68, 0.4); }
.select-status-sc[data-val="Em Cotação"] { background: rgba(59, 130, 246, 0.18); color: #60a5fa; border-color: rgba(59, 130, 246, 0.4); }
.select-status-sc[data-val="Finalizado / Comprado"] { background: rgba(168, 85, 247, 0.18); color: #c084fc; border-color: rgba(168, 85, 247, 0.4); }

.acoes { display: flex; gap: 6px; flex-wrap: wrap; align-items: center; }
.acoes form { margin: 0; }
.acoes button, .acoes .botao { padding: 7px 11px; font-size: 12px; }
.nivel { min-width: 110px; }
.barra { height: 6px; background: #262a36; border-radius: 99px; margin-top: 6px; overflow: hidden; }
.barra span { display: block; height: 100%; background: linear-gradient(90deg, #16a34a, #4ade80); }
.barra.baixo span { background: linear-gradient(90deg, #b91c1c, #f87171); }

/* MODAIS, POP-UP AO VIVO E TOASTS */
.toasts {
    position: fixed; top: 18px; right: 18px; z-index: 9999;
    display: flex; flex-direction: column; gap: 10px; max-width: 390px;
}
.toast {
    background: #14171f; border: 1px solid var(--bd); border-left: 5px solid var(--ok);
    border-radius: 12px; padding: 14px 38px 14px 16px; font-size: 13.5px;
    box-shadow: 0 15px 38px rgba(0, 0, 0, 0.65); position: relative;
    animation: slideInRight .25s ease-out; overflow-wrap: break-word;
}
.toast.erro { border-left-color: var(--er); }
.toast .x {
    position: absolute; right: 8px; top: 6px; background: none;
    color: var(--mt); padding: 4px 8px; font-size: 16px;
}
.overlay {
    display: none; position: fixed; inset: 0; z-index: 9000;
    background: rgba(0, 0, 0, 0.78); backdrop-filter: blur(5px);
    align-items: center; justify-content: center; padding: 16px;
}
.overlay.aberto { display: flex; }
.overlay-alerta-os {
    display: none; position: fixed; inset: 0; z-index: 9800;
    background: rgba(0, 0, 0, 0.74); backdrop-filter: blur(4px);
    align-items: center; justify-content: center; padding: 16px;
}
.overlay-alerta-os.aberto { display: flex; animation: fadeIn .2s ease-out; }
.modal-alerta-os {
    width: 100%; max-width: 520px;
    background: linear-gradient(175deg, #221712, #111318);
    border: 2px solid var(--or); border-radius: 16px; padding: 24px;
    box-shadow: 0 25px 70px rgba(0, 0, 0, 0.9), 0 0 38px rgba(249, 115, 22, 0.35);
}
.alerta-os-topo { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 12px; }
.alerta-os-tag {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 5px 12px; border-radius: 99px; font-size: 12px; font-weight: 800;
    background: rgba(249, 115, 22, 0.22); color: #fed7aa; border: 1px solid var(--or);
}
.alerta-os-grade {
    display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px;
    background: #0c0e12; border: 1px solid var(--bd); border-radius: 12px;
    padding: 14px; margin: 12px 0 16px;
}
.alerta-os-item strong { display: block; font-size: 11px; text-transform: uppercase; color: var(--mt); letter-spacing: .5px; }
.alerta-os-item span { font-size: 13.5px; font-weight: 700; color: var(--tx); overflow-wrap: break-word; }
.alerta-os-item.full { grid-column: 1 / -1; }
.modal {
    width: 100%; max-width: 720px; max-height: 92vh; overflow-y: auto;
    background: linear-gradient(175deg, #1b1f29, #101217);
    border: 1px solid rgba(249, 115, 22, 0.42);
    border-radius: 16px; padding: 24px;
    box-shadow: 0 25px 65px rgba(0, 0, 0, 0.85);
}
.modal h3 { margin: 0 0 6px; font-size: 20px; color: var(--or2); overflow-wrap: break-word; }
.modal p { margin: 0 0 14px; overflow-wrap: break-word; }
.modal .campo { margin-top: 12px; }
.modal-acoes { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; flex-wrap: wrap; }

.stepper { display: flex; gap: 8px; }
.stepper button { width: 44px; padding: 0; font-size: 20px; background: #262a36; flex-shrink: 0; }
.stepper input { text-align: center; font-size: 17px; font-weight: 700; }
.saldo {
    margin-top: 12px !important; padding: 11px; background: #0c0e12;
    border: 1px dashed var(--bd); border-radius: 10px; color: var(--mt);
}
.saldo strong { color: var(--or2); }
.checkbox-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; margin-top: 6px; }
.checkbox-grid label {
    display: flex; align-items: center; gap: 8px; font-weight: 500;
    color: var(--tx); cursor: pointer; background: #0c0e12;
    padding: 9px 11px; border: 1px solid var(--bd); border-radius: 8px; margin: 0;
    line-height: 1.35;
}
.checkbox-grid label:hover { border-color: var(--or); }
.checkbox-grid input[type="checkbox"] { width: auto; accent-color: var(--or); flex-shrink: 0; }
input[type="time"]::-webkit-calendar-picker-indicator,
input[type="date"]::-webkit-calendar-picker-indicator { filter: invert(1); }

/* LOGIN */
.container-login { max-width: 430px; margin: 8vh auto; }
.card-login {
    background: linear-gradient(175deg, var(--s2), var(--s1));
    border: 1px solid var(--bd); border-radius: var(--r); padding: 24px;
}
.card-login .form-login-box { display: grid; gap: 14px; }
.senha-wrap { position: relative; }
.senha-wrap input { padding-right: 48px; }
.campo-senha-mascarada { -webkit-text-security: disc; text-security: disc; }
.campo-senha-mascarada.visivel { -webkit-text-security: none; text-security: none; }
.olho { position: absolute; right: 6px; top: 50%; transform: translateY(-50%); background: none; padding: 6px 10px; }

.bloco-itens {
    border: 1px dashed rgba(249, 115, 22, 0.5); background: #0b0d11;
    border-radius: 12px; padding: 14px; margin-top: 8px;
}
.linha-item-sc {
    display: flex; flex-direction: column; gap: 8px; background: #14171f;
    padding: 12px; border-radius: 8px; border: 1px solid var(--bd); margin-bottom: 10px;
}
.linha-item-sc-baixo { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
.linha-item-sc-baixo input, .linha-item-sc-baixo select { flex: 1; min-width: 120px; }
.rodape { text-align: center; color: #6b7280; font-size: 12px; padding: 16px 0 6px; }

@keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
@keyframes slideInRight { from { opacity: 0; transform: translateX(35px); } to { opacity: 1; transform: translateX(0); } }

/* ============================================================================
   RESPONSIVIDADE MOBILE
   ============================================================================ */
@media (max-width: 860px) {
    body {
        padding: 10px 10px calc(84px + env(safe-area-inset-bottom, 0px));
        line-height: 1.45;
        font-size: 14px;
    }

    .mobile-bottom-nav {
        display: grid; grid-template-columns: repeat(4, 1fr);
        position: fixed; bottom: 0; left: 0; right: 0; z-index: 8500;
        background: rgba(13, 15, 20, 0.96); backdrop-filter: blur(12px);
        border-top: 1px solid rgba(249, 115, 22, 0.35);
        padding: 6px 6px calc(6px + env(safe-area-inset-bottom, 0px));
        box-shadow: 0 -8px 28px rgba(0, 0, 0, 0.75);
    }
    .mobile-bottom-nav a, .mobile-bottom-nav button {
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        gap: 2px; background: transparent; border: 0; border-radius: 10px;
        color: var(--mt); text-decoration: none; font-size: 11px; font-weight: 700;
        padding: 6px 4px; min-height: 46px; box-shadow: none;
    }
    .mobile-bottom-nav a.ativo { color: var(--or2); background: rgba(249, 115, 22, 0.14); }
    .mobile-bottom-nav .nav-ico { font-size: 18px; line-height: 1; }

    .lobby-wrap { margin: 2vh auto; }
    .lobby-card { padding: 20px 14px; border-radius: 14px; }
    .lobby-card h1 { font-size: 19px; }
    .lobby-acoes { flex-direction: column; }
    .btn-entrar-lobby, .lobby-acoes .botao { width: 100%; min-height: 46px; }

    .cabecalho { padding: 14px; gap: 10px; margin-bottom: 14px; border-radius: 12px; }
    .cabecalho-principal { width: 100%; }
    .cabecalho-titulo { align-items: center; gap: 10px; }
    .logo-badge { width: 36px; height: 36px; font-size: 18px; border-radius: 9px; }
    .cabecalho h1 { font-size: 16px; line-height: 1.25; }
    .cabecalho p { font-size: 12px; line-height: 1.35; }
    .nav-abas { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); width: 100%; gap: 7px; margin-top: 12px; }
    .aba, .botao-sair {
        white-space: normal; line-height: 1.25; width: 100%;
        justify-content: center; text-align: center; padding: 10px 8px;
        font-size: 12px; min-height: 42px;
    }
    .cabecalho-lado { width: 100%; flex-direction: row; justify-content: space-between; align-items: center; }
    .relogio { flex: 1; text-align: center; font-size: 11.5px; padding: 7px 10px; }
    .cabecalho-lado .botao-sair { width: auto; padding: 7px 12px; min-height: 36px; }

    .sub-abas-nav {
        display: flex; flex-wrap: nowrap; overflow-x: auto;
        -webkit-overflow-scrolling: touch; gap: 8px; padding: 8px;
        margin-bottom: 14px; scrollbar-width: none;
    }
    .sub-abas-nav::-webkit-scrollbar { display: none; }
    .sub-aba-btn { flex: 0 0 auto; white-space: nowrap; padding: 10px 14px; font-size: 12.5px; min-height: 42px; }

    .resumo { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 9px; margin-bottom: 14px; }
    .resumo-card { padding: 13px; border-radius: 12px; }
    .resumo-card span { font-size: 10.5px; padding-right: 20px; line-height: 1.25; }
    .resumo-card strong { font-size: 22px; line-height: 1.2; }
    .resumo-card small { font-size: 11px; line-height: 1.3; }
    .icone { right: 10px; top: 10px; font-size: 16px; }

    .card { padding: 14px; margin-bottom: 14px; border-radius: 12px; }
    .card h2 { font-size: 15.5px; line-height: 1.3; margin-bottom: 12px; }
    details.card > summary { font-size: 15px; line-height: 1.35; min-height: 36px; }
    details.card > .conteudo { margin-top: 12px; padding-top: 12px; }

    input, textarea, select { font-size: 16px; padding: 11px 12px; min-height: 44px; }
    button, .botao { font-size: 13.5px; padding: 11px 14px; min-height: 44px; }

    .seg { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
    .seg label { min-width: 0; padding: 11px 8px; font-size: 13px; }

    .acoes-topo { flex-direction: column; align-items: stretch; gap: 10px; }
    .acoes-topo .excel, .acoes-topo .imprimir { width: 100%; }
    .busca { min-width: 100%; flex-wrap: nowrap; }
    .busca .limpar { flex-shrink: 0; padding: 10px 14px; }

    .filtros {
        flex-wrap: nowrap; overflow-x: auto; -webkit-overflow-scrolling: touch;
        gap: 8px; margin-top: 12px; padding-bottom: 4px; scrollbar-width: none;
    }
    .filtros::-webkit-scrollbar { display: none; }
    .filtros .rotulo { flex-shrink: 0; }
    .chip { flex-shrink: 0; padding: 8px 13px; font-size: 12.5px; min-height: 38px; }
    .filtros select { flex-shrink: 0; font-size: 14px; min-height: 38px; padding: 6px 12px; }
    .contagem { margin-left: 0; width: 100%; font-size: 12px; margin-top: 8px; display: block; }

    .tabela-wrapper { border: 0; background: transparent; overflow: visible; }
    table.tabela-dinamica { display: block; width: 100%; min-width: 0 !important; }
    table.tabela-dinamica > thead { display: none; }
    table.tabela-dinamica > tbody { display: flex; flex-direction: column; gap: 12px; }
    table.tabela-dinamica > tbody > tr.linha {
        display: flex; flex-direction: column;
        background: linear-gradient(175deg, #191d26, #111318);
        border: 1px solid var(--bd); border-left: 4px solid var(--or);
        border-radius: 12px; padding: 12px 14px;
        box-shadow: 0 6px 18px rgba(0, 0, 0, 0.35);
    }
    table.tabela-dinamica > tbody > tr.linha[data-status="finalizada"] { border-left-color: var(--ok); }
    table.tabela-dinamica > tbody > tr.linha[data-status="aguardando"] { border-left-color: #a855f7; }
    table.tabela-dinamica > tbody > tr.linha > td {
        display: flex; justify-content: space-between; align-items: flex-start;
        gap: 12px; width: 100%; max-width: none !important;
        padding: 6px 0; border-bottom: 1px dashed rgba(255, 255, 255, 0.08);
        font-size: 13.5px; text-align: right;
    }
    table.tabela-dinamica > tbody > tr.linha > td:last-child {
        border-bottom: 0; padding-top: 10px; padding-bottom: 2px; justify-content: stretch;
    }
    table.tabela-dinamica > tbody > tr.linha > td[data-label]::before {
        content: attr(data-label); font-size: 11px; font-weight: 700;
        text-transform: uppercase; color: var(--mt); letter-spacing: 0.4px;
        text-align: left; flex-shrink: 0; padding-top: 2px;
    }
    table.tabela-dinamica > tbody > tr.linha > td:last-child::before { display: none; }
    table.tabela-dinamica > tbody > tr.linha .acoes {
        width: 100%; display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px;
    }
    table.tabela-dinamica > tbody > tr.linha .acoes > form { display: flex; width: 100%; }
    table.tabela-dinamica > tbody > tr.linha .acoes button,
    table.tabela-dinamica > tbody > tr.linha .acoes .botao {
        width: 100%; min-height: 42px; font-size: 12.5px; padding: 8px 10px;
    }
    table.tabela-dinamica > tbody > tr.detalhe { margin-top: -8px; border-radius: 0 0 12px 12px; overflow: hidden; }
    table.tabela-dinamica > tbody > tr.detalhe > td {
        display: block; width: 100%; padding: 14px;
        border-radius: 10px; border: 1px solid var(--bd); border-left: 4px solid var(--or);
    }
    table.tabela-dinamica > tbody > tr.sem-resultado > td,
    table.tabela-dinamica > tbody > tr > td.vazio { display: block; width: 100%; }
    .galeria-evidencias { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    .select-status-sc { width: 100%; max-width: 190px; font-size: 13px; min-height: 38px; }
    .item-arquivo-acoes { width: 100%; display: grid; grid-template-columns: 1fr auto; }
    .item-arquivo-acoes > form { display: flex; }

    table:not(.tabela-dinamica) { min-width: 560px; }

    .overlay, .overlay-alerta-os { padding: 0; align-items: flex-end; }
    .modal, .modal-alerta-os {
        padding: 20px 16px calc(22px + env(safe-area-inset-bottom, 0px));
        border-radius: 18px 18px 0 0; max-height: 92vh; overflow-y: auto; border-bottom: 0;
    }
    .modal h3 { font-size: 17.5px; line-height: 1.3; }
    .modal-acoes { flex-direction: column-reverse; gap: 8px; }
    .modal-acoes button, .modal-acoes .botao { width: 100%; min-height: 46px; }
    .alerta-os-grade { grid-template-columns: 1fr; }

    .form-grid, .form-grid.tres, .checkbox-grid, .grid-duplo, .grid-triplo { grid-template-columns: 1fr; gap: 12px; }
    .campo.completo { grid-column: auto; }
    .toasts { left: 10px; right: 10px; top: 10px; max-width: none; }
    .toast { font-size: 13.5px; line-height: 1.4; }
}
"""

JS_BASE = r"""
(function(){
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('/sw.js').catch(()=>{});
}

window.addEventListener('load', () => {
    const btnInstalar = document.getElementById('btnInstalarApp');
    if (btnInstalar) {
        btnInstalar.style.display = 'inline-flex';
        btnInstalar.addEventListener('click', async () => {
            if (window.deferredPrompt) {
                window.deferredPrompt.prompt();
                const { outcome } = await window.deferredPrompt.userChoice;
                if (outcome === 'accepted') btnInstalar.style.display = 'none';
                window.deferredPrompt = null;
            } else {
                alert('Para instalar no celular:\n\n• Android (Chrome): toque nos 3 pontinhos e escolha "Adicionar à tela inicial".\n• iPhone (Safari): toque em Compartilhar e escolha "Adicionar à Tela de Início".');
            }
        });
    }
});

window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    window.deferredPrompt = e;
    const btnInstalar = document.getElementById('btnInstalarApp');
    if (btnInstalar) btnInstalar.style.display = 'inline-flex';
});

const $=(s,r)=>(r||document).querySelector(s), $$=(s,r)=>Array.from((r||document).querySelectorAll(s));
const norm=t=>(t||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');

const caixa=document.createElement('div');caixa.className='toasts';document.body.appendChild(caixa);
window.toast=function(msg,tipo){
    const t=document.createElement('div');
    t.className='toast '+(tipo||'ok');
    const s=document.createElement('span');s.textContent=msg;
    const x=document.createElement('button');x.type='button';x.className='x';x.textContent='×';
    x.onclick=()=>t.remove();
    t.append(s,x);caixa.appendChild(t);
    setTimeout(()=>t.remove(),5000);
};
(window.__FLASH__||[]).forEach((m,i)=>setTimeout(()=>window.toast(m,/erro|não|nao|inválid|insuficiente|incorret|obrigat|excede|limite/i.test(m)?'erro':'ok'),i*220));

/* ============================================================================
   SOM DE NOTIFICAÇÃO E POP-UP AO VIVO DE NOVA O.S.
   ============================================================================ */
let audioCtx = null;
function desbloquearAudio(){
    try {
        const Ctx = window.AudioContext || window.webkitAudioContext;
        if (!Ctx) return;
        if (!audioCtx) audioCtx = new Ctx();
        if (audioCtx.state === 'suspended') {
            audioCtx.resume().catch(()=>{});
        }
    } catch(_e){}
}
['click', 'keydown', 'touchstart'].forEach(ev => {
    document.addEventListener(ev, desbloquearAudio, { passive: true });
});

window.tocarSomNotificacaoOS = function(){
    try {
        const Ctx = window.AudioContext || window.webkitAudioContext;
        if (!Ctx) return;
        if (!audioCtx) audioCtx = new Ctx();
        if (audioCtx.state === 'suspended') {
            audioCtx.resume().catch(()=>{});
        }
        const agora = audioCtx.currentTime;
        const notas = [
            { freq: 880, inicio: 0.00, duracao: 0.14 },
            { freq: 1174.66, inicio: 0.16, duracao: 0.22 }
        ];
        notas.forEach(n => {
            const osc = audioCtx.createOscillator();
            const gain = audioCtx.createGain();
            osc.type = 'triangle';
            osc.frequency.setValueAtTime(n.freq, agora + n.inicio);
            gain.gain.setValueAtTime(0.0001, agora + n.inicio);
            gain.gain.exponentialRampToValueAtTime(0.24, agora + n.inicio + 0.02);
            gain.gain.exponentialRampToValueAtTime(0.0001, agora + n.inicio + n.duracao);
            osc.connect(gain);
            gain.connect(audioCtx.destination);
            osc.start(agora + n.inicio);
            osc.stop(agora + n.inicio + n.duracao + 0.02);
        });
    } catch(_e){}
};

const ovAlertaOS = document.createElement('div');
ovAlertaOS.className = 'overlay-alerta-os';
ovAlertaOS.id = 'overlayAlertaNovaOS';
ovAlertaOS.innerHTML = `
    <div class="modal-alerta-os" role="dialog" aria-modal="true" aria-labelledby="alertaOSTitulo">
        <div class="alerta-os-topo">
            <span class="alerta-os-tag">🔔 NOVA O.S. REGISTRADA</span>
            <button type="button" class="limpar" id="btnFecharAlertaOS" style="padding:6px 12px;font-size:12px;min-height:36px">✕ Fechar</button>
        </div>
        <h3 id="alertaOSTitulo" style="margin:0 0 4px;font-size:20px;color:#fb923c">Ordem de Serviço #<span id="alertaOSNumero"></span></h3>
        <p class="muted" style="margin:0">Registrada em <strong style="color:#f3f4f6" id="alertaOSAbertura"></strong></p>
        <div class="alerta-os-grade">
            <div class="alerta-os-item"><strong>Máquina / Equipamento</strong><span id="alertaOSSetor">-</span></div>
            <div class="alerta-os-item"><strong>Prioridade</strong><span id="alertaOSPrioridade">-</span></div>
            <div class="alerta-os-item"><strong>Solicitante</strong><span id="alertaOSSolicitante">-</span></div>
            <div class="alerta-os-item"><strong>Setor</strong><span id="alertaOSResponsavel">-</span></div>
            <div class="alerta-os-item full"><strong>Causa / Sintoma</strong><span id="alertaOSTituloCausa">-</span></div>
        </div>
        <div class="modal-acoes" style="margin-top:12px">
            <button type="button" class="limpar" id="btnOkAlertaOS">Entendido</button>
            <a href="/os" class="botao salvar" id="btnVerNaFilaOS">📋 Ver na Fila de O.S.</a>
        </div>
    </div>
`;
document.body.appendChild(ovAlertaOS);

const filaAlertasOS = [];
let alertaOSAtivo = false;

function fecharAlertaOSAtual(){
    ovAlertaOS.classList.remove('aberto');
    alertaOSAtivo = false;
    if (filaAlertasOS.length > 0) {
        setTimeout(exibirProximoAlertaOS, 220);
    }
}

function exibirProximoAlertaOS(){
    if (alertaOSAtivo || filaAlertasOS.length === 0) return;
    const os = filaAlertasOS.shift();
    alertaOSAtivo = true;

    $('#alertaOSNumero').textContent = os.numero_os || '';
    $('#alertaOSAbertura').textContent = os.abertura || '';
    $('#alertaOSSetor').textContent = os.setor || 'Não informado';
    $('#alertaOSPrioridade').textContent = os.prioridade || 'Média';
    $('#alertaOSSolicitante').textContent = os.solicitante || '-';
    $('#alertaOSResponsavel').textContent = os.responsavel || '-';
    $('#alertaOSTituloCausa').textContent = os.titulo || 'Sem causa informada';

    const btnVer = $('#btnVerNaFilaOS');
    btnVer.href = '/os?busca=' + encodeURIComponent(os.numero_os || '');

    ovAlertaOS.classList.add('aberto');
    window.tocarSomNotificacaoOS();
}

$('#btnFecharAlertaOS')?.addEventListener('click', fecharAlertaOSAtual);
$('#btnOkAlertaOS')?.addEventListener('click', fecharAlertaOSAtual);
ovAlertaOS.addEventListener('mousedown', e => {
    if (e.target === ovAlertaOS) fecharAlertaOSAtual();
});

window.dispararPopupNovaOS = function(os){
    if (!os || !os.id) return;
    filaAlertasOS.push(os);
    exibirProximoAlertaOS();
};

const CHAVE_ULTIMO_OS = 'jamaica_ultimo_os_id_visto';
function lerUltimoIdVisto(){
    const v = parseInt(localStorage.getItem(CHAVE_ULTIMO_OS) || '0', 10);
    return isNaN(v) ? 0 : v;
}
function salvarUltimoIdVisto(id){
    const atual = lerUltimoIdVisto();
    if (id > atual) {
        localStorage.setItem(CHAVE_ULTIMO_OS, String(id));
    }
}

function iniciarMonitoramentoNovaOS(){
    if (typeof window.__ULTIMO_OS_ID__ === 'number' && window.__ULTIMO_OS_ID__ > 0) {
        salvarUltimoIdVisto(window.__ULTIMO_OS_ID__);
    }

    async function verificarNovasOS(){
        if (document.hidden) return;
        const ultimoId = lerUltimoIdVisto();
        try {
            const resp = await fetch(`/api/os/novas?ultimo_id=${ultimoId}`, {
                cache: 'no-store',
                headers: { 'X-Requested-With': 'XMLHttpRequest' }
            });
            if (!resp.ok) return;
            const data = await resp.json();
            if (!data || !Array.isArray(data.ordens)) return;
            if (data.ordens.length > 0) {
                let maiorId = ultimoId;
                data.ordens.forEach(os => {
                    if (os.id > maiorId) maiorId = os.id;
                    if (os.id > ultimoId) {
                        window.dispararPopupNovaOS(os);
                    }
                });
                salvarUltimoIdVisto(maiorId);
            }
        } catch(_e){}
    }

    setInterval(verificarNovasOS, 3500);
    document.addEventListener('visibilitychange', () => {
        if (!document.hidden) verificarNovasOS();
    });
}
iniciarMonitoramentoNovaOS();

window.confirmar=function(msg,rotulo,perigo){
    return new Promise(resolve=>{
        const ov=document.createElement('div');ov.className='overlay aberto';
        const box=document.createElement('div');box.className='modal';
        const h=document.createElement('h3');h.textContent='Confirmar ação';
        const p=document.createElement('p');p.textContent=msg;
        const a=document.createElement('div');a.className='modal-acoes';
        const n=document.createElement('button');n.type='button';n.className='limpar';n.textContent='Cancelar';
        const y=document.createElement('button');y.type='button';y.className=perigo?'excluir':'salvar';y.textContent=rotulo||'Confirmar';
        a.append(n,y);box.append(h,p,a);ov.appendChild(box);document.body.appendChild(ov);
        const end=v=>{ov.remove();resolve(v)};
        n.onclick=()=>end(false);
        y.onclick=()=>end(true);
        ov.onclick=e=>{if(e.target===ov)end(false)};
        document.addEventListener('keydown',function esc(e){if(e.key==='Escape'){document.removeEventListener('keydown',esc);end(false)}});
    });
};
document.addEventListener('submit',e=>{
    const f=e.target;
    if(!f.dataset||!f.dataset.confirm||f.dataset.ok)return;
    e.preventDefault();
    window.confirmar(f.dataset.confirm,f.dataset.confirmBtn,f.dataset.perigo==='1').then(ok=>{
        if(ok){f.dataset.ok='1';f.submit();}
    });
});

const rel=$('#relogio');
if(rel){
    const tic=()=>rel.textContent=new Date().toLocaleString('pt-BR',{weekday:'short',day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit'});
    tic();setInterval(tic,1000);
}

$$('textarea[data-max]').forEach(ta=>{
    const c=document.createElement('div');
    c.style.cssText='text-align:right;color:#9ca3af;font-size:11px;margin-top:4px';
    ta.insertAdjacentElement('afterend',c);
    ta.maxLength=+ta.dataset.max;
    const up=()=>c.textContent=ta.value.length+' / '+ta.dataset.max;
    ta.addEventListener('input',up);up();
});

function initTabelas(){
    $$('table.tabela-dinamica').forEach(tabela=>{
        const tbody=tabela.tBodies[0];if(!tbody)return;
        const cabecalhos=$$('thead th',tabela).map(th=>th.textContent.trim());
        const linhas=$$('tr.linha',tbody),vazio=$('tr.sem-resultado',tbody),contador=$('#contador-'+tabela.id);
        const estado={q:'',f:{}};
        linhas.forEach(tr=>{
            $$('td',tr).forEach((td,idx)=>{
                if(!td.dataset.label && cabecalhos[idx]) td.dataset.label=cabecalhos[idx];
            });
            tr._det=tr.nextElementSibling&&tr.nextElementSibling.classList.contains('detalhe')?tr.nextElementSibling:null;
            tr._txt=norm(tr.textContent+' '+(tr._det?tr._det.textContent:''));
        });
        const exibirDetalhe = (det, aberto) => {
            if (!det) return;
            if (!aberto) { det.style.display = 'none'; return; }
            det.style.display = window.innerWidth <= 860 ? 'block' : 'table-row';
        };
        function aplicar(){
            let n=0;
            linhas.forEach(tr=>{
                let ok=!estado.q||tr._txt.includes(estado.q);
                Object.keys(estado.f).forEach(k=>{if(estado.f[k]&&tr.dataset[k]!==estado.f[k])ok=false});
                tr.style.display=ok?'':'none';
                if(tr._det) exibirDetalhe(tr._det, ok && tr._det.classList.contains('aberto'));
                if(ok)n++;
            });
            if(vazio)vazio.style.display=n===0&&linhas.length?(window.innerWidth<=860?'block':'table-row'):'none';
            if(contador)contador.innerHTML='Exibindo <strong>'+n+'</strong> de '+linhas.length+' registros';
        }
        if(tabela.classList.contains('detalhavel')){
            linhas.forEach(tr=>{
                if(!tr._det)return;
                tr.classList.add('clicavel');
                tr.addEventListener('click',e=>{
                    if(e.target.closest('button,a,form,input,select'))return;
                    tr._det.classList.toggle('aberto');
                    exibirDetalhe(tr._det, tr._det.classList.contains('aberto'));
                });
            });
        }
        $$('th[data-tipo]',tabela).forEach(th=>{
            th.classList.add('ordenavel');
            th.addEventListener('click',()=>{
                const asc=th.dataset.dir!=='asc';
                $$('th[data-tipo]',tabela).forEach(x=>delete x.dataset.dir);
                th.dataset.dir=asc?'asc':'desc';
                const i=th.cellIndex,tipo=th.dataset.tipo;
                const val=tr=>{
                    const t=tr.cells[i].textContent.trim();
                    if(!t||t==='-')return null;
                    if(tipo==='num'){const n=parseFloat(t);return isNaN(n)?null:n}
                    if(tipo==='data'){const m=t.match(/(\d+)\/(\d+)\/(\d+)\s+(\d+):(\d+):(\d+)/);return m?new Date(m[3],m[2]-1,m[1],m[4],m[5],m[6]).getTime():null}
                    return norm(t);
                };
                linhas.sort((a,b)=>{
                    const x=val(a),y=val(b);
                    if(x===null)return 1;if(y===null)return -1;
                    const c=typeof x==='number'?x-y:x.localeCompare(y,'pt-BR',{numeric:true});
                    return asc?c:-c;
                }).forEach(tr=>{tbody.appendChild(tr);if(tr._det)tbody.appendChild(tr._det)});
                aplicar();
            });
        });
        window.addEventListener('resize', aplicar);
        aplicar();
        const wrapper=tabela.closest('.painel-secao')||document;
        const q=$('.q-busca',wrapper);
        if(q){
            q.addEventListener('input',()=>{estado.q=norm(q.value.trim());aplicar()});
            if(q.value){estado.q=norm(q.value);aplicar()}
        }
        $$('[data-filtro]',wrapper).forEach(el=>{
            if(el.tagName==='SELECT')el.addEventListener('change',()=>{estado.f[el.dataset.filtro]=el.value;aplicar()});
            else el.addEventListener('click',()=>{
                $$('[data-filtro="'+el.dataset.filtro+'"]',wrapper).forEach(x=>x.classList.remove('ativo'));
                el.classList.add('ativo');estado.f[el.dataset.filtro]=el.dataset.valor;aplicar();
            });
        });
        $$('[data-atalho]',wrapper).forEach(c=>c.addEventListener('click',()=>{
            const a=$(c.dataset.atalho,wrapper);if(a){a.click();wrapper.scrollIntoView({behavior:'smooth'})}
        }));
    });
}
initTabelas();

$$('.sub-aba-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
        $$('.sub-aba-btn').forEach(b=>b.classList.remove('ativa'));
        $$('.painel-secao').forEach(p=>p.classList.remove('ativo'));
        btn.classList.add('ativa');
        const alvo=$('#'+btn.dataset.alvo);
        if(alvo) alvo.classList.add('ativo');
        localStorage.setItem('gestao_aba_ativa', btn.dataset.alvo);
    });
});
const salvaAba=localStorage.getItem('gestao_aba_ativa');
if(salvaAba && $('#'+salvaAba)){
    $('[data-alvo="'+salvaAba+'"]')?.click();
}

$$('form').forEach(f=>{
    if((f.method||'').toLowerCase()!=='post')return;
    if($('input[name="csrf_token"]',f))return;
    const h=document.createElement('input');h.type='hidden';h.name='csrf_token';h.value=window.__CSRF__||'';f.appendChild(h);
});

function vincularPreviewMultiplasImagens(inputEl, boxEl, gridEl, resumoEl){
    let urlsAtuais = [];
    const limpar = () => {
        if (inputEl) inputEl.value = '';
        urlsAtuais.forEach(u => URL.revokeObjectURL(u));
        urlsAtuais = [];
        if (gridEl) gridEl.innerHTML = '';
        if (resumoEl) resumoEl.textContent = '';
        if (boxEl) boxEl.style.display = 'none';
    };
    if (inputEl) {
        inputEl.addEventListener('change', () => {
            urlsAtuais.forEach(u => URL.revokeObjectURL(u));
            urlsAtuais = [];
            if (gridEl) gridEl.innerHTML = '';

            const arquivos = inputEl.files ? Array.from(inputEl.files) : [];
            if (!arquivos.length) {
                if (boxEl) boxEl.style.display = 'none';
                return;
            }
            if (resumoEl) {
                resumoEl.textContent = arquivos.length === 1
                    ? '1 foto selecionada:'
                    : `${arquivos.length} fotos selecionadas:`;
            }
            arquivos.forEach(arq => {
                const url = URL.createObjectURL(arq);
                urlsAtuais.push(url);
                if (gridEl) {
                    const card = document.createElement('div');
                    card.className = 'preview-grid-item';
                    const img = document.createElement('img');
                    img.src = url;
                    img.alt = arq.name;
                    const cap = document.createElement('span');
                    cap.textContent = arq.name;
                    cap.title = arq.name;
                    card.append(img, cap);
                    gridEl.appendChild(card);
                }
            });
            if (boxEl) boxEl.style.display = 'block';
        });
    }
    return limpar;
}

/* MODAL EDITAR O.S. REGISTRADA */
const modalEditarOS = $('#modalEditarOS');
if(modalEditarOS){
    const formEditarOS = $('#formEditarOS');
    $$('.btn-editar-os').forEach(btn => {
        btn.addEventListener('click', () => {
            formEditarOS.action = btn.dataset.url;
            $('#osNumeroEditar').textContent = btn.dataset.numero || '';
            $('#editOsSolicitante').value = btn.dataset.solicitante || '';
            $('#editOsSetor').value = btn.dataset.setor || '';
            $('#editOsTitulo').value = btn.dataset.titulo || '';
            $('#editOsResponsavel').value = btn.dataset.responsavel || '';
            $('#editOsPrioridade').value = btn.dataset.prioridade || 'Média';
            $('#editOsStatus').value = btn.dataset.status || 'Aberta';
            $('#editOsTipo').value = btn.dataset.tipo || '';
            $('#editOsTempo').value = btn.dataset.tempo || '01:00';
            $('#editOsObs').value = btn.dataset.obs || '';

            const selecionados = (btn.dataset.funcionarios || '')
                .split(',')
                .map(s => s.trim())
                .filter(Boolean);
            $$('input[name="funcionarios"]', formEditarOS).forEach(cb => {
                cb.checked = selecionados.includes(cb.value) || (
                    selecionados.includes('Jefferson') && cb.value.startsWith('Jefferson')
                );
            });

            modalEditarOS.classList.add('aberto');
        });
    });
    const fecharEditOS = () => modalEditarOS.classList.remove('aberto');
    $$('[data-fechar-edit-os]', modalEditarOS).forEach(b => b.addEventListener('click', fecharEditOS));
    modalEditarOS.addEventListener('mousedown', e => { if(e.target === modalEditarOS) fecharEditOS(); });
}

const modalFinalizar = $('#modalFinalizar');
if(modalFinalizar){
    const formFinalizar = $('#formFinalizar');
    const limparPreviewEvidencia = vincularPreviewMultiplasImagens(
        $('#evidenciaFotoInput'),
        $('#previewEvidenciaBox'),
        $('#previewEvidenciaGrid'),
        $('#previewEvidenciaResumo')
    );

    $$('.btn-abrir-finalizar').forEach(btn => {
        btn.addEventListener('click', () => {
            formFinalizar.action = btn.dataset.url;
            $('#osNumeroFinalizar').textContent = btn.dataset.numero;
            $$('input[name="funcionarios"]', formFinalizar).forEach(cb => cb.checked = false);
            $('#tempoTrabalhado').value = '01:00';
            if($('#obsTecnica')) $('#obsTecnica').value = '';
            if($('#pecaUtilizada')) $('#pecaUtilizada').value = '';
            if($('#qtdPecaUtilizada')) $('#qtdPecaUtilizada').value = '1';
            limparPreviewEvidencia();
            modalFinalizar.classList.add('aberto');
        });
    });
    const fecharFin = () => {
        limparPreviewEvidencia();
        modalFinalizar.classList.remove('aberto');
    };
    $$('[data-fechar-fin]', modalFinalizar).forEach(b => b.addEventListener('click', fecharFin));
    modalFinalizar.addEventListener('mousedown', e => { if(e.target === modalFinalizar) fecharFin(); });
}

const modalEvidenciaOS = $('#modalEvidenciaOS');
if(modalEvidenciaOS){
    const formEvidenciaOS = $('#formEvidenciaOS');
    const formRemoverTodasEv = $('#formRemoverEvidenciaOS');
    const boxAtual = $('#evidenciaAtualModalBox');
    const galeriaAtual = $('#galeriaEvidenciasModal');
    const limparPreviewModalEv = vincularPreviewMultiplasImagens(
        $('#evidenciaModalInput'),
        $('#previewModalEvBox'),
        $('#previewModalEvGrid'),
        $('#previewModalEvResumo')
    );

    $$('.btn-gerenciar-evidencia').forEach(btn => {
        btn.addEventListener('click', () => {
            formEvidenciaOS.action = btn.dataset.url;
            $('#osNumeroEvidencia').textContent = btn.dataset.numero || '';
            limparPreviewModalEv();
            galeriaAtual.innerHTML = '';

            let lista = [];
            try {
                lista = JSON.parse(btn.dataset.evidencias || '[]');
            } catch(_e) {
                lista = [];
            }

            if (Array.isArray(lista) && lista.length > 0) {
                boxAtual.style.display = 'block';
                const osId = btn.dataset.osId || '';
                lista.forEach(item => {
                    const card = document.createElement('div');
                    card.className = 'evidencia-item-card';

                    const link = document.createElement('a');
                    link.href = item.url;
                    link.target = '_blank';
                    link.rel = 'noopener noreferrer';

                    const img = document.createElement('img');
                    img.src = item.url;
                    img.alt = item.nome || item.arquivo;
                    link.appendChild(img);

                    const cap = document.createElement('span');
                    cap.className = 'evidencia-item-nome';
                    cap.textContent = item.nome || item.arquivo;
                    cap.title = item.nome || item.arquivo;

                    const fRm = document.createElement('form');
                    fRm.method = 'POST';
                    fRm.action = `/os/evidencia-remover-item/${osId}`;
                    fRm.style.margin = '0';

                    const hCsrf = document.createElement('input');
                    hCsrf.type = 'hidden';
                    hCsrf.name = 'csrf_token';
                    hCsrf.value = window.__CSRF__ || '';

                    const hArq = document.createElement('input');
                    hArq.type = 'hidden';
                    hArq.name = 'arquivo';
                    hArq.value = item.arquivo;

                    const bRm = document.createElement('button');
                    bRm.type = 'submit';
                    bRm.className = 'excluir';
                    bRm.style.cssText = 'width:100%;padding:5px 8px;font-size:11px;min-height:32px';
                    bRm.textContent = '✕ Remover';

                    fRm.append(hCsrf, hArq, bRm);
                    card.append(link, cap, fRm);
                    galeriaAtual.appendChild(card);
                });

                if (formRemoverTodasEv && btn.dataset.removeUrl) {
                    formRemoverTodasEv.action = btn.dataset.removeUrl;
                    formRemoverTodasEv.style.display = 'inline-flex';
                }
            } else {
                boxAtual.style.display = 'none';
                if (formRemoverTodasEv) {
                    formRemoverTodasEv.style.display = 'none';
                }
            }
            modalEvidenciaOS.classList.add('aberto');
        });
    });

    const fecharModalEv = () => {
        limparPreviewModalEv();
        modalEvidenciaOS.classList.remove('aberto');
    };
    $$('[data-fechar-evidencia]', modalEvidenciaOS).forEach(b => b.addEventListener('click', fecharModalEv));
    modalEvidenciaOS.addEventListener('mousedown', e => { if(e.target === modalEvidenciaOS) fecharModalEv(); });
}

const modalPausar = $('#modalPausar');
if(modalPausar){
    const formPausar = $('#formPausar');
    $$('.btn-abrir-pausar').forEach(btn => {
        btn.addEventListener('click', () => {
            formPausar.action = btn.dataset.url;
            $('#osNumeroPausar').textContent = btn.dataset.numero;
            $('#motivoPausaInput').value = btn.dataset.motivo || '';
            modalPausar.classList.add('aberto');
        });
    });
    const fecharPausa = () => modalPausar.classList.remove('aberto');
    $$('[data-fechar-pausa]', modalPausar).forEach(b => b.addEventListener('click', fecharPausa));
    modalPausar.addEventListener('mousedown', e => { if(e.target === modalPausar) fecharPausa(); });
}

const modalMov = $('#modalMov');
if(modalMov){
    const formMov = $('#formMov'), movQtd = $('#movQtd'), movSaldo = $('#movSaldo'), movSol = $('#movSol'), movLblSol = $('#movLblSol');
    let qtdAtual = 0;
    const calc = () => {
        const q = parseInt(movQtd.value, 10) || 0;
        const tipo = $('input[name="tipo"]:checked', formMov).value;
        const saldo = tipo === 'Retirada' ? qtdAtual - q : qtdAtual + q;
        movSaldo.textContent = saldo;
        movSaldo.style.color = saldo < 0 ? '#f87171' : '#fb923c';
        if(tipo === 'Retirada'){ movSol.setAttribute('required', 'required'); movLblSol.textContent = 'Solicitante / O.S. *'; }
        else { movSol.removeAttribute('required'); movLblSol.textContent = 'Responsável / Nota Fiscal'; }
    };
    $$('.movimentar').forEach(btn => {
        btn.addEventListener('click', () => {
            formMov.action = btn.dataset.url;
            $('#movTitulo').textContent = '⇄ Movimentar: ' + btn.dataset.nome;
            $('#movSub').textContent = 'Código: ' + btn.dataset.codigo + ' | Saldo Atual: ' + btn.dataset.qtd;
            qtdAtual = parseInt(btn.dataset.qtd, 10) || 0;
            movQtd.value = 1;
            movSol.value = '';
            $('#tipoRet').checked = true;
            calc();
            modalMov.classList.add('aberto');
        });
    });
    $$('input[name="tipo"]', formMov).forEach(r => r.addEventListener('change', calc));
    movQtd.addEventListener('input', calc);
    $$('.stepper button', modalMov).forEach(b => {
        b.addEventListener('click', () => {
            const step = parseInt(b.dataset.step, 10) || 1;
            movQtd.value = Math.max(1, (parseInt(movQtd.value, 10) || 0) + step);
            calc();
        });
    });
    const fecharMov = () => modalMov.classList.remove('aberto');
    $$('[data-fechar]', modalMov).forEach(b => b.addEventListener('click', fecharMov));
    modalMov.addEventListener('mousedown', e => { if(e.target === modalMov) fecharMov(); });
}

const modalEditarPeca = $('#modalEditarPeca');
if(modalEditarPeca){
    const formEditarPeca = $('#formEditarPeca');
    $$('.editar-peca-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            formEditarPeca.action = btn.dataset.url;
            $('#editarCodigoPeca').value = btn.dataset.codigo || '';
            $('#editarNomePeca').value = btn.dataset.nome || '';
            $('#editarCategoriaPeca').value = btn.dataset.categoria || '';
            $('#editarLocalizacaoPeca').value = btn.dataset.localizacao || '';
            $('#editarEstoqueMinimoPeca').value = btn.dataset.estoqueMinimo || '2';
            modalEditarPeca.classList.add('aberto');
        });
    });
    const fecharEdicao = () => modalEditarPeca.classList.remove('aberto');
    $('#cancelarEdicaoPeca')?.addEventListener('click', fecharEdicao);
    modalEditarPeca.addEventListener('mousedown', e => { if(e.target === modalEditarPeca) fecharEdicao(); });
}

const modalEditarSpare = $('#modalEditarSpare');
if(modalEditarSpare){
    const formEditarSpare = $('#formEditarSpare');
    $$('.btn-editar-spare').forEach(btn => {
        btn.addEventListener('click', () => {
            formEditarSpare.action = btn.dataset.url;
            $('#spareMaquinaInput').value = btn.dataset.maquina || '';
            $('#spareModeloInput').value = btn.dataset.modelo || '';
            $('#spareSetorInput').value = btn.dataset.setor || '';
            $('#spareNotasInput').value = btn.dataset.notas || '';
            modalEditarSpare.classList.add('aberto');
        });
    });
    const fecharSpare = () => modalEditarSpare.classList.remove('aberto');
    $('#cancelarEdicaoSpare')?.addEventListener('click', fecharSpare);
    modalEditarSpare.addEventListener('mousedown', e => { if(e.target === modalEditarSpare) fecharSpare(); });
}
})();
"""

# ==============================================================================
# TEMPLATES HTML
# ==============================================================================
HTML_LOBBY = r"""
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><link rel="icon" type="image/svg+xml" href="{{favicon}}">{{pwa_head|safe}}<title>LOBBY — MANUTENÇÃO JAMAICA</title><style>{{css|safe}}</style></head><body>
<div class="lobby-wrap">
    <div class="lobby-card">
        <div class="lobby-logo">J</div>
        <h1>CENTRAL DE MANUTENÇÃO — JAMAICA</h1>
        <p class="sub">Ordens de Serviço, Estoque, Spare Parts, PCM e Solicitações de Compra.</p>

        <div style="display:flex;justify-content:center;margin-bottom:20px">
            <span class="relogio" id="relogio"></span>
        </div>

        <div class="resumo" style="margin-bottom:8px;text-align:left">
            <div class="resumo-card"><span class="icone">📋</span><span>Total de O.S.</span><strong>{{total}}</strong><small>Histórico geral</small></div>
            <div class="resumo-card"><span class="icone">🟡</span><span>O.S. Abertas</span><strong style="color:#fde047">{{abertas}}</strong><small>Em fila</small></div>
            <div class="resumo-card"><span class="icone">⏳</span><span>Aguard. Peça</span><strong style="color:#d8b4fe">{{aguardando}}</strong><small>Suprimentos</small></div>
            <div class="resumo-card"><span class="icone">🔢</span><span>Próxima O.S.</span><strong>{{proxima}}</strong><small>Numeração auto</small></div>
        </div>

        <div class="lobby-acoes">
            <a href="{{url_for('dashboard')}}" class="botao btn-entrar-lobby">🚀 Entrar no Sistema</a>
            <a href="{{url_for('central_gestao')}}" class="botao limpar" style="padding:14px 20px;border-radius:12px">🔐 Central de Gestão</a>
            <button type="button" id="btnInstalarApp" class="botao limpar" style="display:inline-flex;padding:14px 20px;border-radius:12px">📲 Instalar App</button>
        </div>
    </div>
    <footer class="rodape">Central de Manutenção JAMAICA</footer>
</div>
<nav class="mobile-bottom-nav" aria-label="Navegação Rápida Mobile">
    <a href="{{url_for('lobby')}}" class="ativo"><span class="nav-ico">🏠</span>Lobby</a>
    <a href="{{url_for('dashboard')}}"><span class="nav-ico">📋</span>O.S.</a>
    <a href="{{url_for('central_gestao')}}"><span class="nav-ico">🔐</span>Gestão</a>
    <a href="{{url_for('dashboard')}}#cardNovaOS"><span class="nav-ico">➕</span>Nova O.S.</a>
</nav>
<script>window.__FLASH__={{flash_msgs|tojson}};window.__ULTIMO_OS_ID__={{ultimo_os_id|default(0)}};</script><script>{{js|safe}}</script></body></html>
"""

HTML_OS = r"""
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><link rel="icon" type="image/svg+xml" href="{{favicon}}">{{pwa_head|safe}}<title>CENTRAL O.S. — JAMAICA</title><style>{{css|safe}}</style></head><body><div class="container">
<header class="cabecalho">
    <div class="cabecalho-principal">
        <div class="cabecalho-titulo">
            <div class="logo-badge">J</div>
            <div class="cabecalho-texto">
                <h1>CENTRAL DE ORDENS DE SERVIÇO — JAMAICA</h1>
                <p>Gestão de manutenção industrial e abertura em chão de fábrica</p>
            </div>
        </div>
        <div class="nav-abas">
            <a href="{{url_for('lobby')}}" class="aba">🏠 Lobby</a>
            <a href="{{url_for('dashboard')}}" class="aba ativa">📋 Ordens de Serviço</a>
            <a href="{{url_for('central_gestao')}}" class="aba" style="border-color:var(--or2);color:#fed7aa">🔐 Central de Gestão</a>
            <button type="button" id="btnInstalarApp" class="aba" style="display:inline-flex;background:var(--btn);border-color:var(--or);color:#fff">📲 Instalar App</button>
        </div>
    </div>
    <div class="cabecalho-lado"><span class="relogio" id="relogio"></span></div>
</header>

<section class="resumo">
    <div class="resumo-card" data-atalho='[data-filtro="status"][data-valor=""]'><span class="icone">📋</span><span>Total de O.S.</span><strong>{{total}}</strong><div class="progresso"><span style="width:{{pct_final}}%"></span></div><small>{{pct_final}}% concluídas</small></div>
    <div class="resumo-card" data-atalho='[data-filtro="status"][data-valor="aberta"]'><span class="icone">🟡</span><span>O.S. Abertas</span><strong style="color:#fde047">{{abertas}}</strong><small>Em fila / execução</small></div>
    <div class="resumo-card" data-atalho='[data-filtro="status"][data-valor="aguardando"]'><span class="icone">⏳</span><span>Aguardando Peça</span><strong style="color:#d8b4fe">{{aguardando}}</strong><small>Suprimentos</small></div>
    <div class="resumo-card" data-atalho='[data-filtro="status"][data-valor="finalizada"]'><span class="icone">✅</span><span>O.S. Finalizadas</span><strong style="color:#4ade80">{{finalizadas}}</strong><small>Próxima O.S.: <b>{{proxima}}</b></small></div>
</section>

<details class="card" id="cardNovaOS" open><summary>➕ Abrir Nova Ordem de Serviço</summary><div class="conteudo"><form id="formCriarOS" action="{{url_for('criar')}}" method="POST"><div class="form-grid">
    <div class="campo"><label>Número da O.S.</label><input value="{{proxima}}" disabled></div>
    <div class="campo"><label>Solicitante *</label><input name="solicitante" maxlength="100" required placeholder="Nome do solicitante"></div>
    <div class="campo"><label>Máquina / Equipamento *</label><input name="setor" list="listaMaquinas" maxlength="100" value="{{prefill_setor}}" placeholder="Ex: Extrusora Troester 90/16">
        <datalist id="listaMaquinas">{% for m in maquinas_sugeridas %}<option value="{{m}}">{% endfor %}</datalist>
    </div>
    <div class="campo"><label>Causa / Sintoma</label><input name="titulo" maxlength="200" placeholder="Resumo objetivo da falha ou solicitação"></div>
    <div class="campo"><label>Prioridade</label>
        <div class="seg">
            <input type="radio" name="prioridade" id="p1" value="Baixa"><label for="p1">Baixa</label>
            <input type="radio" name="prioridade" id="p2" value="Média" checked><label for="p2">Média</label>
            <input type="radio" name="prioridade" id="p3" value="Alta"><label for="p3">Alta</label>
            <input type="radio" name="prioridade" id="p4" value="Crítica"><label for="p4">Crítica</label>
        </div>
    </div>
    <div class="campo"><label>Setor</label><input name="responsavel" maxlength="100" placeholder="Local aonde há o problema"></div>
    <div class="campo completo"><button class="salvar" style="width:100%" type="submit" onclick="window.tocarSomNotificacaoOS && window.tocarSomNotificacaoOS()">+ Registrar Ordem de Serviço</button></div>
</div></form></div></details>

<section class="card" id="secao-tabela">
    <div class="acoes-topo">
        <form class="busca" action="{{url_for('dashboard')}}" method="GET">
            <div class="campo-busca"><input id="q" class="q-busca" name="busca" value="{{busca}}" autocomplete="off" placeholder="Pesquisar por número, solicitante, máquina..."></div>
            <a class="botao limpar" href="{{url_for('dashboard')}}">Limpar</a>
        </form>
        <a class="botao excel" href="{{url_for('excel')}}">⬇ Exportar Excel</a>
    </div>
    <div class="filtros">
        <span class="rotulo">Status</span>
        <button type="button" class="chip ativo" data-filtro="status" data-valor="">Todas</button>
        <button type="button" class="chip" data-filtro="status" data-valor="aberta">Abertas</button>
        <button type="button" class="chip" data-filtro="status" data-valor="aguardando">Aguard. Peça</button>
        <button type="button" class="chip" data-filtro="status" data-valor="finalizada">Finalizadas</button>
        <span class="rotulo">Prioridade</span>
        <select data-filtro="prioridade">
            <option value="">Todas</option>
            <option value="critica">Crítica</option>
            <option value="alta">Alta</option>
            <option value="media">Média</option>
            <option value="baixa">Baixa</option>
        </select>
        <span class="contagem" id="contador-tabelaOS"></span>
    </div>
</section>

<section class="card">
    <h2>Ordens de Serviço Registradas</h2>
    <div class="tabela-wrapper">
        <table id="tabelaOS" class="tabela-dinamica detalhavel">
            <thead>
                <tr>
                    <th data-tipo="num">OS</th>
                    <th data-tipo="texto">Solicitante</th>
                    <th data-tipo="texto">Máquina / Setor</th>
                    <th data-tipo="texto">Causa</th>
                    <th data-tipo="texto">Prioridade</th>
                    <th data-tipo="texto">Responsável</th>
                    <th data-tipo="texto">Tipo</th>
                    <th data-tipo="texto">Status</th>
                    <th data-tipo="data">Abertura</th>
                    <th data-tipo="data">Finalização</th>
                    <th>Lead Time</th>
                    <th>Ações</th>
                </tr>
            </thead>
            <tbody>
            {% for os in ordens %}
                {% set st_code = 'finalizada' if os.status=='Finalizada' else ('aguardando' if os.status=='Aguardando Peça' else 'aberta') %}
                {% set lista_ev = os.evidencias_lista %}
                <tr class="linha" data-status="{{st_code}}" data-prioridade="{{os.prioridade|lower|replace('í','i')|replace('é','e')}}">
                    <td data-label="OS"><strong>{{os.numero_os}}{% if lista_ev %} 📷 ({{lista_ev|length}}){% endif %}</strong></td>
                    <td data-label="Solicitante">{{os.solicitante}}</td>
                    <td data-label="Máquina / Setor" style="max-width:190px">{{os.setor or '-'}}</td>
                    <td data-label="Causa" style="max-width:240px">{{os.titulo or '-'}}</td>
                    <td data-label="Prioridade"><span class="prioridade prioridade-{{os.prioridade|lower|replace('í','i')|replace('é','e')}}">{{os.prioridade}}</span></td>
                    <td data-label="Responsável">{{os.responsavel or '-'}}</td>
                    <td data-label="Tipo">{% if os.tipo_manutencao %}<span class="tipo-badge tipo-{{os.tipo_manutencao|lower}}">{{os.tipo_manutencao}}</span>{% else %}<span class="muted">-</span>{% endif %}</td>
                    <td data-label="Status">
                        {% if os.status=='Finalizada' %}<span class="status finalizada">Finalizada</span>
                        {% elif os.status=='Aguardando Peça' %}<span class="status aguardando-peca">⏳ Aguard. Peça</span>
                        {% else %}<span class="status aberta">Aberta</span>{% endif %}
                    </td>
                    <td data-label="Abertura">{{os.abertura}}</td>
                    <td data-label="Finalização">{{os.finalizacao or '-'}}</td>
                    <td data-label="Lead Time">{{os.tempo_total or '-'}}</td>
                    <td data-label="Ações">
                        <div class="acoes">
                            <button type="button" class="editar-peca btn-editar-os"
                                data-url="{{url_for('editar_os', id=os.id)}}"
                                data-numero="{{os.numero_os}}"
                                data-solicitante="{{os.solicitante}}"
                                data-setor="{{os.setor or ''}}"
                                data-titulo="{{os.titulo or ''}}"
                                data-responsavel="{{os.responsavel or ''}}"
                                data-prioridade="{{os.prioridade}}"
                                data-status="{{os.status}}"
                                data-tipo="{{os.tipo_manutencao or ''}}"
                                data-tempo="{{fmt_input_time(os.horas_trabalhadas)}}"
                                data-funcionarios="{{os.funcionarios or ''}}"
                                data-obs="{{os.observacao_tecnica or ''}}"
                                title="Editar O.S.">✏️ Editar</button>
                            {% if os.status!='Finalizada' %}
                            <button type="button" class="finalizar btn-abrir-finalizar" data-url="{{url_for('finalizar',id=os.id)}}" data-numero="{{os.numero_os}}">✔ Finalizar</button>
                            <button type="button" class="pausar btn-abrir-pausar" data-url="{{url_for('pausar_os',id=os.id)}}" data-numero="{{os.numero_os}}" data-motivo="{{os.motivo_pausa}}" title="Pausar por falta de peça ou retomar">⏳ Espera</button>
                            {% else %}
                            <button type="button" class="evidencia-btn btn-gerenciar-evidencia"
                                data-os-id="{{os.id}}"
                                data-url="{{url_for('atualizar_evidencia_os', id=os.id)}}"
                                data-remove-url="{{url_for('remover_evidencia_rota', id=os.id)}}"
                                data-numero="{{os.numero_os}}"
                                data-evidencias='{{lista_ev|tojson}}'
                                title="Anexar ou gerenciar fotos de evidência">📷 {{ 'Fotos (' ~ (lista_ev|length) ~ ')' if lista_ev else '+ Evidências' }}</button>
                            {% endif %}
                            <a href="{{url_for('imprimir_os',id=os.id)}}" target="_blank" class="botao imprimir" title="Imprimir O.S. em A4">🖨️ A4</a>
                            <form action="{{url_for('excluir',id=os.id)}}" method="POST" data-perigo="1" data-confirm="Tem certeza que deseja excluir permanentemente a O.S. {{os.numero_os}}?" data-confirm-btn="Excluir">
                                <button class="excluir" title="Excluir O.S.">✕ Excluir</button>
                            </form>
                        </div>
                    </td>
                </tr>
                <tr class="detalhe">
                    <td colspan="12">
                        {% if os.motivo_pausa and os.status=='Aguardando Peça' %}
                        <p><strong style="color:#d8b4fe">⏳ Motivo da Pausa (Suprimentos):</strong> {{os.motivo_pausa}}</p>
                        {% endif %}
                        {% if os.observacao_tecnica %}
                        <p><strong style="color:#4ade80">Laudo / Serviço Executado:</strong> {{os.observacao_tecnica}}</p>
                        {% endif %}
                        {% if os.funcionarios %}
                        <p><strong>Executantes:</strong> {{os.funcionarios}} &nbsp;|&nbsp; <strong>Horas Apontadas:</strong> {{fmt_hhmm(os.horas_trabalhadas)}}</p>
                        {% endif %}
                        {% if not os.observacao_tecnica and not os.funcionarios and not os.motivo_pausa and not lista_ev %}
                        <p class="muted">Sem observações ou anexos adicionais nesta O.S.</p>
                        {% endif %}
                        {% if lista_ev %}
                        <div class="evidencia-box">
                            <strong style="color:#60a5fa">📷 Evidências do Trabalho Feito ({{lista_ev|length}}):</strong>
                            <div class="galeria-evidencias">
                                {% for ev in lista_ev %}
                                <div class="evidencia-item-card">
                                    <a href="{{ev.url}}" target="_blank" rel="noopener noreferrer" title="Abrir {{ev.nome}}">
                                        <img src="{{ev.url}}" alt="Evidência {{loop.index}} da O.S. {{os.numero_os}}">
                                    </a>
                                    <span class="evidencia-item-nome" title="{{ev.nome}}">{{loop.index}}. {{ev.nome}}</span>
                                </div>
                                {% endfor %}
                            </div>
                        </div>
                        {% endif %}
                    </td>
                </tr>
            {% else %}
                <tr><td colspan="12" class="vazio">Nenhuma Ordem de Serviço encontrada.</td></tr>
            {% endfor %}
                <tr class="sem-resultado"><td colspan="12" class="vazio">Nenhum resultado corresponde aos filtros aplicados.</td></tr>
            </tbody>
        </table>
    </div>
</section>
<footer class="rodape">Central de Manutenção JAMAICA</footer></div>

<nav class="mobile-bottom-nav" aria-label="Navegação Rápida Mobile">
    <a href="{{url_for('lobby')}}"><span class="nav-ico">🏠</span>Lobby</a>
    <a href="{{url_for('dashboard')}}" class="ativo"><span class="nav-ico">📋</span>O.S.</a>
    <a href="{{url_for('central_gestao')}}"><span class="nav-ico">🔐</span>Gestão</a>
    <a href="#cardNovaOS" onclick="document.getElementById('cardNovaOS').open=true"><span class="nav-ico">➕</span>Nova O.S.</a>
</nav>

<!-- MODAL EDITAR O.S. -->
<div class="overlay" id="modalEditarOS">
    <div class="modal">
        <h3>✏️ Editar Ordem de Serviço <span id="osNumeroEditar"></span></h3>
        <form id="formEditarOS" method="POST">
            <div class="form-grid">
                <div class="campo">
                    <label>Solicitante *</label>
                    <input type="text" name="solicitante" id="editOsSolicitante" maxlength="100" required>
                </div>
                <div class="campo">
                    <label>Máquina / Equipamento</label>
                    <input type="text" name="setor" id="editOsSetor" list="listaMaquinas" maxlength="100">
                </div>
                <div class="campo">
                    <label>Causa / Sintoma</label>
                    <input type="text" name="titulo" id="editOsTitulo" maxlength="200">
                </div>
                <div class="campo">
                    <label>Setor</label>
                    <input type="text" name="responsavel" id="editOsResponsavel" maxlength="100">
                </div>
                <div class="campo">
                    <label>Prioridade</label>
                    <select name="prioridade" id="editOsPrioridade">
                        <option value="Baixa">Baixa</option>
                        <option value="Média">Média</option>
                        <option value="Alta">Alta</option>
                        <option value="Crítica">Crítica</option>
                    </select>
                </div>
                <div class="campo">
                    <label>Status</label>
                    <select name="status" id="editOsStatus">
                        <option value="Aberta">Aberta</option>
                        <option value="Aguardando Peça">Aguardando Peça</option>
                        <option value="Finalizada">Finalizada</option>
                    </select>
                </div>
                <div class="campo">
                    <label>Tipo de Manutenção</label>
                    <select name="tipo_manutencao" id="editOsTipo">
                        <option value="">-- Não definido --</option>
                        <option value="Preventiva">Preventiva</option>
                        <option value="Corretiva">Corretiva</option>
                    </select>
                </div>
                <div class="campo">
                    <label>Horas Trabalhadas (HH:MM)</label>
                    <input type="time" name="tempo_trabalhado" id="editOsTempo">
                </div>
                <div class="campo completo">
                    <label>Funcionários Envolvidos</label>
                    <div class="checkbox-grid">
                        {% for f in funcionarios_equipe %}
                        <label><input type="checkbox" name="funcionarios" value="{{f.valor}}"> {{f.icone}} {{f.nome}} ({{f.funcao}} • {{f.turno}})</label>
                        {% endfor %}
                    </div>
                </div>
                <div class="campo completo">
                    <label>Laudo Técnico / Observação</label>
                    <textarea name="observacao_tecnica" id="editOsObs" rows="3"></textarea>
                </div>
            </div>
            <div class="modal-acoes">
                <button type="button" class="limpar" data-fechar-edit-os>Cancelar</button>
                <button type="submit" class="salvar">Salvar Alterações</button>
            </div>
        </form>
    </div>
</div>

<div class="overlay" id="modalFinalizar">
    <div class="modal">
        <h3>🔧 Finalizar Ordem de Serviço <span id="osNumeroFinalizar"></span></h3>
        <form id="formFinalizar" method="POST" enctype="multipart/form-data">
            <div class="campo">
                <label>Tipo de Manutenção *</label>
                <div class="seg">
                    <input type="radio" id="tipoPrev" name="tipo_manutencao" value="Preventiva" checked>
                    <label for="tipoPrev">🛡️ Preventiva</label>
                    <input type="radio" id="tipoCorr" name="tipo_manutencao" value="Corretiva">
                    <label for="tipoCorr">⚡ Corretiva</label>
                </div>
            </div>
            <div class="campo">
                <label>Funcionários Envolvidos *</label>
                <div class="checkbox-grid">
                    {% for f in funcionarios_equipe %}
                    <label><input type="checkbox" name="funcionarios" value="{{f.valor}}"> {{f.icone}} {{f.nome}} ({{f.funcao}} • {{f.turno}})</label>
                    {% endfor %}
                </div>
            </div>
            <div class="campo">
                <label>Horas Trabalhadas (HH:MM) *</label>
                <input type="time" name="tempo_trabalhado" id="tempoTrabalhado" required value="01:00">
            </div>
            <div class="campo">
                <label>Laudo Técnico / Serviço Realizado</label>
                <textarea name="observacao_tecnica" id="obsTecnica" rows="3" placeholder="Descreva o que foi reparado, ajustado ou substituído..."></textarea>
            </div>
            <div class="campo">
                <label>📷 Evidências do Trabalho Feito (PNG, JPG, WEBP até 6 MB cada)</label>
                <input type="file" name="evidencia_foto" id="evidenciaFotoInput" accept="image/png,image/jpeg,image/webp" multiple>
                <div class="preview-evidencia-box" id="previewEvidenciaBox">
                    <span class="muted" id="previewEvidenciaResumo"></span>
                    <div class="preview-grid" id="previewEvidenciaGrid"></div>
                </div>
            </div>
            {% if pecas_estoque %}
            <div class="form-grid" style="margin-top:14px">
                <div class="campo">
                    <label>Baixar Peça do Estoque (Opcional)</label>
                    <select name="peca_id" id="pecaUtilizada">
                        <option value="">-- Nenhuma peça utilizada --</option>
                        {% for p in pecas_estoque %}
                        <option value="{{p.id}}">{{p.codigo}} — {{p.nome}} (Saldo: {{p.quantidade}})</option>
                        {% endfor %}
                    </select>
                </div>
                <div class="campo">
                    <label>Quantidade Utilizada</label>
                    <input type="number" name="peca_qtd" id="qtdPecaUtilizada" min="1" value="1">
                </div>
            </div>
            {% endif %}
            <div class="modal-acoes">
                <button type="button" class="limpar" data-fechar-fin>Cancelar</button>
                <button type="submit" class="salvar">Confirmar e Finalizar</button>
            </div>
        </form>
    </div>
</div>

<div class="overlay" id="modalEvidenciaOS">
    <div class="modal">
        <h3>📷 Galeria de Evidências — O.S. <span id="osNumeroEvidencia"></span></h3>
        <div class="evidencia-box" id="evidenciaAtualModalBox" style="display:none;margin-bottom:14px">
            <strong style="color:#4ade80;font-size:12.5px">Fotos Anexadas:</strong>
            <div class="galeria-evidencias" id="galeriaEvidenciasModal"></div>
        </div>

        <form id="formEvidenciaOS" method="POST" enctype="multipart/form-data">
            <div class="campo">
                <label>Adicionar Novas Fotos *</label>
                <input type="file" name="evidencia_foto" id="evidenciaModalInput" accept="image/png,image/jpeg,image/webp" multiple required>
                <div class="preview-evidencia-box" id="previewModalEvBox">
                    <span class="muted" id="previewModalEvResumo"></span>
                    <div class="preview-grid" id="previewModalEvGrid"></div>
                </div>
            </div>
            <div class="modal-acoes">
                <button type="button" class="limpar" data-fechar-evidencia>Fechar</button>
                <button type="submit" class="salvar">💾 Enviar Fotos</button>
            </div>
        </form>

        <form id="formRemoverEvidenciaOS" method="POST" style="display:none;margin-top:10px;width:100%" data-perigo="1" data-confirm="Deseja remover permanentemente TODAS as fotos de evidência desta O.S.?" data-confirm-btn="Remover Todas">
            <button type="submit" class="excluir" style="width:100%">🗑️ Remover Todas as Fotos</button>
        </form>
    </div>
</div>

<div class="overlay" id="modalPausar">
    <div class="modal">
        <h3>⏳ Status de Espera — O.S. <span id="osNumeroPausar"></span></h3>
        <form id="formPausar" method="POST">
            <div class="campo">
                <label>Status Operacional</label>
                <div class="seg">
                    <input type="radio" id="stAguard" name="novo_status" value="Aguardando Peça" checked>
                    <label for="stAguard">⏳ Aguardando Peça</label>
                    <input type="radio" id="stAberta" name="novo_status" value="Aberta">
                    <label for="stAberta">🟡 Retomar (Aberta)</label>
                </div>
            </div>
            <div class="campo">
                <label>Observação / Número da S.C.</label>
                <input type="text" name="motivo_pausa" id="motivoPausaInput" maxlength="200" placeholder="Ex: Aguardando rolamento 6204 — S.C. 0496">
            </div>
            <div class="modal-acoes">
                <button type="button" class="limpar" data-fechar-pausa>Cancelar</button>
                <button type="submit" class="salvar">Salvar Status</button>
            </div>
        </form>
    </div>
</div>
<script>window.__FLASH__={{flash_msgs|tojson}};window.__ULTIMO_OS_ID__={{ultimo_os_id|default(0)}};</script><script>{{js|safe}}</script></body></html>
"""

HTML_LOGIN_SISTEMA = r"""
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><link rel="icon" type="image/svg+xml" href="{{favicon}}">{{pwa_head|safe}}<title>LOGIN — CENTRAL GESTÃO JAMAICA</title><style>{{css|safe}}</style></head><body class="login">
<div class="container-login">
    <div class="cabecalho" style="justify-content:center;text-align:center">
        <div class="cabecalho-principal">
            <h1>🔒 ÁREA RESTRITA</h1>
            <p>Digite as credenciais para acessar a Central de Gestão</p>
        </div>
    </div>
    <div class="card-login">
        <div class="form-login-box" id="boxLoginAcesso">
            <div class="campo">
                <label>Usuário</label>
                <input type="text" id="campoUsuarioLogin" required autofocus autocomplete="off" autocapitalize="none" autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" placeholder="Digite o usuário">
            </div>
            <div class="campo">
                <label>Senha</label>
                <div class="senha-wrap">
                    <input type="text" id="campoSenhaLogin" class="campo-senha-mascarada" required autocomplete="off" autocapitalize="none" autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" placeholder="Digite a senha">
                    <button type="button" class="olho" id="verSenhaLogin" title="Mostrar/Ocultar senha">👁️</button>
                </div>
            </div>
            <button class="salvar" type="button" id="btnEntrarGestao">Entrar na Central</button>
            <a href="{{url_for('dashboard')}}" class="botao limpar">Ir para Ordens de Serviço</a>
            <a href="{{url_for('lobby')}}" class="botao limpar">Voltar ao Lobby</a>
        </div>
    </div>
</div>
<script>
window.__FLASH__={{flash_msgs|tojson}};
window.__ULTIMO_OS_ID__={{ultimo_os_id|default(0)}};
document.addEventListener('DOMContentLoaded', function(){
    const inpUser = document.getElementById('campoUsuarioLogin');
    const inpPass = document.getElementById('campoSenhaLogin');
    const btnOlho = document.getElementById('verSenhaLogin');
    const btnEntrar = document.getElementById('btnEntrarGestao');

    if (btnOlho && inpPass) {
        btnOlho.addEventListener('click', function(){
            inpPass.classList.toggle('visivel');
        });
    }

    async function autenticarSemPopup(){
        const u = (inpUser.value || '').trim();
        const s = (inpPass.value || '').trim();
        if (!u || !s) {
            window.toast('Preencha usuário e senha.', 'erro');
            return;
        }
        btnEntrar.disabled = true;
        btnEntrar.textContent = 'Validando...';
        try {
            const resp = await fetch("{{url_for('login_sistema')}}", {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-Requested-With': 'XMLHttpRequest',
                    'X-CSRF-Token': window.__CSRF__ || ''
                },
                body: JSON.stringify({ usuario: u, senha: s })
            });
            const data = await resp.json();
            if (resp.ok && data.ok) {
                inpUser.value = '';
                inpPass.value = '';
                window.location.replace(data.redirect || "{{url_for('central_gestao')}}");
                return;
            }
            window.toast(data.erro || 'Usuário ou senha incorretos.', 'erro');
        } catch (_e) {
            window.toast('Erro de conexão ao validar acesso.', 'erro');
        } finally {
            btnEntrar.disabled = false;
            btnEntrar.textContent = 'Entrar na Central';
        }
    }

    btnEntrar?.addEventListener('click', autenticarSemPopup);
    [inpUser, inpPass].forEach(el => {
        el?.addEventListener('keydown', function(e){
            if (e.key === 'Enter') {
                e.preventDefault();
                autenticarSemPopup();
            }
        });
    });
});
</script>
<script>{{js|safe}}</script></body></html>
"""

HTML_CENTRAL_GESTAO = r"""
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><link rel="icon" type="image/svg+xml" href="{{favicon}}">{{pwa_head|safe}}<title>CENTRAL DE GESTÃO — JAMAICA</title><style>{{css|safe}}</style>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head><body><div class="container">
<header class="cabecalho">
    <div class="cabecalho-principal">
        <div class="cabecalho-titulo">
            <div class="logo-badge">J</div>
            <div class="cabecalho-texto">
                <h1>CENTRAL DE GESTÃO DE MANUTENÇÃO — JAMAICA</h1>
                <p>Estoque, Spare Parts, PCM, Compras, Indicadores & Governança</p>
            </div>
        </div>
        <div class="nav-abas">
            <a href="{{url_for('lobby')}}" class="aba">🏠 Lobby</a>
            <a href="{{url_for('dashboard')}}" class="aba">📋 Ordens de Serviço</a>
            <a href="{{url_for('central_gestao')}}" class="aba ativa">🔐 Central de Gestão</a>
            <button type="button" id="btnInstalarApp" class="aba" style="display:inline-flex;background:var(--btn);border-color:var(--or);color:#fff">📲 Instalar App</button>
        </div>
    </div>
    <div class="cabecalho-lado"><span class="relogio" id="relogio"></span><a href="{{url_for('logout_sistema')}}" class="botao-sair">Sair</a></div>
</header>

<div class="sub-abas-nav">
    <button type="button" class="sub-aba-btn ativa" data-alvo="secao-estoque">📦 Estoque</button>
    <button type="button" class="sub-aba-btn" data-alvo="secao-spare">🔧 Spare Parts & QR</button>
    <button type="button" class="sub-aba-btn" data-alvo="secao-pcm">📅 PCM</button>
    <button type="button" class="sub-aba-btn" data-alvo="secao-compras">🛒 Compras (S.C.)</button>
    <button type="button" class="sub-aba-btn" data-alvo="secao-indicadores">📊 Indicadores & Equipe</button>
    <button type="button" class="sub-aba-btn" data-alvo="secao-governanca">🛡️ Backup & Auditoria</button>
</div>

<!-- ================= 1. ESTOQUE ================= -->
<div id="secao-estoque" class="painel-secao ativo">
    <section class="resumo">
        <div class="resumo-card"><span class="icone">🔩</span><span>SKUs Cadastrados</span><strong>{{total_pecas}}</strong></div>
        <div class="resumo-card"><span class="icone">📦</span><span>Volume Total</span><strong>{{total_quantidade}}</strong></div>
        <div class="resumo-card" data-atalho='[data-filtro="baixo"][data-valor="1"]'><span class="icone">🚨</span><span>Estoque Baixo</span><strong style="color:#f87171">{{itens_baixo}}</strong><small>Filtrar baixos</small></div>
        <div class="resumo-card"><span class="icone">🔁</span><span>Movimentações</span><strong>{{total_movimentacoes}}</strong></div>
    </section>

    <details class="card"><summary>➕ Cadastrar Nova Peça no Estoque</summary><div class="conteudo"><form action="{{url_for('cadastrar_peca')}}" method="POST"><div class="form-grid tres">
        <div class="campo"><label>Código / SKU *</label><input name="codigo" required maxlength="50" placeholder="Ex: ROL-6204" style="text-transform:uppercase"></div>
        <div class="campo"><label>Nome da Peça *</label><input name="nome" required maxlength="150" placeholder="Ex: Rolamento Rígido de Esferas 6204"></div>
        <div class="campo"><label>Categoria</label><input name="categoria" placeholder="Ex: Mecânica, Elétrica, Pneumática"></div>
        <div class="campo"><label>Localização Física</label><input name="localizacao" placeholder="Ex: Corredor A / Prateleira B-3"></div>
        <div class="campo"><label>Quantidade Inicial *</label><input type="number" name="quantidade" min="0" value="1" required></div>
        <div class="campo"><label>Estoque Mínimo *</label><input type="number" name="estoque_minimo" min="1" value="2" required></div>
        <div class="campo completo"><button class="salvar" style="width:100%" type="submit">+ Cadastrar Peça</button></div>
    </div></form></div></details>

    <section class="card">
        <div class="acoes-topo">
            <div class="busca"><div class="campo-busca"><input class="q-busca" placeholder="Pesquisar peça no estoque..."></div></div>
            <a class="botao excel" href="{{url_for('excel_estoque')}}">⬇ Exportar Excel (Estoque)</a>
        </div>
        <div class="filtros">
            <span class="rotulo">Exibir</span>
            <button type="button" class="chip ativo" data-filtro="baixo" data-valor="">Todas as peças</button>
            <button type="button" class="chip" data-filtro="baixo" data-valor="1">🚨 Somente estoque baixo</button>
            <span class="contagem" id="contador-tabelaEstoque"></span>
        </div>
    </section>

    <section class="card">
        <h2>Peças em Estoque</h2>
        <div class="tabela-wrapper">
            <table id="tabelaEstoque" class="tabela-dinamica">
                <thead>
                    <tr><th data-tipo="texto">Código</th><th data-tipo="texto">Peça</th><th data-tipo="texto">Categoria</th><th data-tipo="texto">Localização</th><th data-tipo="num">Saldo</th><th data-tipo="num">Mín.</th><th>Status</th><th data-tipo="data">Última Atualização</th><th>Ações</th></tr>
                </thead>
                <tbody>
                {% for p in pecas %}
                    {% set baixo=p.quantidade<=p.estoque_minimo %}
                    {% set pct=[((p.quantidade/((p.estoque_minimo*3) or 1))*100)|round|int,100]|min %}
                    <tr class="linha" data-baixo="{{'1' if baixo else ''}}">
                        <td data-label="Código"><strong>{{p.codigo}}</strong></td>
                        <td data-label="Peça">{{p.nome}}</td>
                        <td data-label="Categoria">{{p.categoria or '-'}}</td>
                        <td data-label="Localização">{{p.localizacao or '-'}}</td>
                        <td data-label="Saldo" class="nivel"><strong>{{p.quantidade}}</strong><div class="barra {{'baixo' if baixo else ''}}"><span style="width:{{pct}}%"></span></div></td>
                        <td data-label="Mínimo">{{p.estoque_minimo}}</td>
                        <td data-label="Status">{% if baixo %}<span class="status estoque-baixo">Baixo Estoque</span>{% else %}<span class="status estoque-ok">Normal</span>{% endif %}</td>
                        <td data-label="Atualização">{{p.ultima_atualizacao}}</td>
                        <td data-label="Ações">
                            <div class="acoes">
                                <button type="button" class="movimentar" data-url="{{url_for('movimentar_estoque',id=p.id)}}" data-nome="{{p.nome}}" data-codigo="{{p.codigo}}" data-qtd="{{p.quantidade}}" data-min="{{p.estoque_minimo}}">⇄ Movimentar</button>
                                {% if baixo %}
                                <a href="{{url_for('central_gestao', repor_codigo=p.codigo, repor_nome=p.nome, repor_qtd=(p.estoque_minimo*2 - p.quantidade) if (p.estoque_minimo*2 - p.quantidade) > 0 else p.estoque_minimo, tab='compras')}}" class="botao imprimir" title="Gerar S.C. de Reposição">🛒 Pedir S.C.</a>
                                {% endif %}
                                <button type="button" class="editar-peca-btn editar-peca" data-url="{{url_for('editar_peca',id=p.id)}}" data-codigo="{{p.codigo}}" data-nome="{{p.nome}}" data-categoria="{{p.categoria}}" data-localizacao="{{p.localizacao}}" data-estoque-minimo="{{p.estoque_minimo}}">✏️ Editar</button>
                                <form action="{{url_for('excluir_peca',id=p.id)}}" method="POST" data-perigo="1" data-confirm="Excluir peça {{p.nome}}?" data-confirm-btn="Excluir"><button class="excluir">✕ Excluir</button></form>
                            </div>
                        </td>
                    </tr>
                {% else %}
                    <tr><td colspan="9" class="vazio">Nenhuma peça cadastrada.</td></tr>
                {% endfor %}
                    <tr class="sem-resultado"><td colspan="9" class="vazio">Nenhum resultado.</td></tr>
                </tbody>
            </table>
        </div>
    </section>

    <section class="card">
        <h2>Histórico Recente de Movimentações</h2>
        <div class="tabela-wrapper">
            <table style="min-width:680px">
                <thead><tr><th>Data / Hora</th><th>Tipo</th><th>Código</th><th>Peça</th><th>Qtd</th><th>Solicitante / Origem</th></tr></thead>
                <tbody>
                {% for m in movimentacoes %}
                    <tr>
                        <td>{{m.data_hora}}</td>
                        <td>{% if m.tipo=='Retirada' %}<span style="color:#f59e0b;font-weight:700">▼ Retirada</span>{% else %}<span style="color:#34d399;font-weight:700">▲ Entrada</span>{% endif %}</td>
                        <td><strong>{{m.peca_codigo}}</strong></td>
                        <td>{{m.peca_nome}}</td>
                        <td><strong>{{m.quantidade}}</strong></td>
                        <td>{{m.solicitante or '-'}}</td>
                    </tr>
                {% else %}
                    <tr><td colspan="6" class="vazio">Nenhuma movimentação.</td></tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
    </section>
</div>

<!-- ================= 2. SPARE PARTS ================= -->
<div id="secao-spare" class="painel-secao">
    <details class="card" open><summary>➕ Cadastrar Máquina / Equipamento (Spare Parts)</summary><div class="conteudo">
    <form action="{{url_for('cadastrar_maquina_spare')}}" method="POST"><div class="form-grid tres">
        <div class="campo"><label>Máquina / Equipamento *</label><input name="maquina" required maxlength="150" placeholder="Ex: Misturador Banbury"></div>
        <div class="campo"><label>Modelo / Fabricante *</label><input name="modelo" required maxlength="100" placeholder="Ex: Alawik 40L"></div>
        <div class="campo"><label>Setor</label><input name="setor" maxlength="100" placeholder="Ex: Massa"></div>
        <div class="campo completo"><label>Observações / Peças Críticas Recomendadas</label><input name="pecas_necessarias" placeholder="Ex: Kit vedação, rolamentos..."></div>
        <div class="campo completo"><button class="salvar" style="width:100%" type="submit">+ Cadastrar Equipamento</button></div>
    </div></form></div></details>

    <section class="card">
        <div class="acoes-topo">
            <div class="busca"><div class="campo-busca"><input class="q-busca" placeholder="Pesquisar máquina..."></div></div>
            <a href="{{url_for('etiquetas_qrcode')}}" target="_blank" class="botao imprimir">📱 Imprimir Etiquetas QR Code</a>
        </div>
    </section>

    <section class="card">
    <h2>Relação de Máquinas e Spare Parts</h2>
    <div class="tabela-wrapper">
    <table id="tabelaSpare" class="tabela-dinamica detalhavel">
    <thead><tr><th data-tipo="texto">Máquina</th><th data-tipo="texto">Modelo</th><th data-tipo="texto">Setor</th><th>Notas</th><th>Peças Mapeadas</th><th>Ações</th></tr></thead>
    <tbody>
    {% for m in maquinas %}
    <tr class="linha">
        <td data-label="Máquina"><strong>{{m.maquina}}</strong></td>
        <td data-label="Modelo">{{m.modelo}}</td>
        <td data-label="Setor">{{m.setor or '-'}}</td>
        <td data-label="Notas">{{m.pecas_necessarias or '-'}}</td>
        <td data-label="Peças"><span class="status estoque-ok">{{m.pecas|length}} item(ns)</span></td>
        <td data-label="Ações">
            <div class="acoes">
                <a href="{{url_for('dashboard', abrir_maquina=m.maquina)}}" class="botao limpar" title="Abrir O.S.">+ Abrir O.S.</a>
                <button type="button" class="editar-peca btn-editar-spare"
                    data-url="{{url_for('editar_maquina_spare', id=m.id)}}"
                    data-maquina="{{m.maquina}}"
                    data-modelo="{{m.modelo}}"
                    data-setor="{{m.setor}}"
                    data-notas="{{m.pecas_necessarias}}"
                    title="Editar Máquina e Notas">✏️ Editar</button>
                <form action="{{url_for('excluir_maquina_spare', id=m.id)}}" method="POST" data-perigo="1" data-confirm="Excluir equipamento {{m.maquina}}?" data-confirm-btn="Excluir">
                    <button class="excluir">✕ Excluir</button>
                </form>
            </div>
        </td>
    </tr>
    <tr class="detalhe">
        <td colspan="6">
            <div style="padding:6px 0">
                <h4 style="margin:0 0 10px;color:var(--or2)">⚙️ Lista Técnica de Reposição</h4>
                {% if m.pecas %}
                <div class="tabela-wrapper" style="overflow-x:auto;margin-bottom:14px">
                <table style="background:#14171f;min-width:620px">
                    <thead><tr><th>Componente / Peça</th><th>Código</th><th>Qtd</th><th>Saldo Estoque</th><th>Obs</th><th>Ação</th></tr></thead>
                    <tbody>
                    {% for p in m.pecas %}
                    {% set saldo_est = mapa_estoque.get(p.codigo_peca|upper) if p.codigo_peca else None %}
                    <tr>
                        <td><strong>{{p.nome_peca}}</strong></td>
                        <td>{{p.codigo_peca or '-'}}</td>
                        <td>{{p.quantidade}}</td>
                        <td>
                            {% if saldo_est is not none %}
                                {% if saldo_est >= p.quantidade %}<span class="status estoque-ok">✔ {{saldo_est}} em estoque</span>
                                {% elif saldo_est > 0 %}<span class="status aberta">⚠️ {{saldo_est}} (Baixo)</span>
                                {% else %}<span class="status estoque-baixo">🚨 0 (Zerado)</span>{% endif %}
                            {% else %}<span class="muted">SKU não vinculado</span>{% endif %}
                        </td>
                        <td>{{p.observacao or '-'}}</td>
                        <td>
                            <form action="{{url_for('excluir_peca_spare', id=p.id)}}" method="POST"><button class="excluir" style="padding:4px 8px;font-size:11px">✕</button></form>
                        </td>
                    </tr>
                    {% endfor %}
                    </tbody>
                </table>
                </div>
                {% else %}
                <p class="muted" style="margin-bottom:12px">Nenhuma peça vinculada.</p>
                {% endif %}
                <details style="background:#111318;padding:12px;border-radius:10px;border:1px solid var(--bd)">
                    <summary style="cursor:pointer;color:var(--or2);font-weight:700">+ Vincular Peça de Reposição</summary>
                    <form action="{{url_for('adicionar_peca_spare', id=m.id)}}" method="POST" style="margin-top:12px">
                        <div class="form-grid tres">
                            <div class="campo"><label>Nome da Peça *</label><input name="nome_peca" required placeholder="Ex: Retificador"></div>
                            <div class="campo"><label>Código / SKU</label><input name="codigo_peca" placeholder="Ex: RET-01" style="text-transform:uppercase"></div>
                            <div class="campo"><label>Qtd</label><input type="number" name="quantidade" value="1" min="1"></div>
                            <div class="campo completo"><label>Observação</label><input name="observacao" placeholder="Aplicação..."></div>
                            <div class="campo completo"><button class="salvar" style="width:100%" type="submit">+ Salvar Peça</button></div>
                        </div>
                    </form>
                </details>
            </div>
        </td>
    </tr>
    {% else %}
    <tr><td colspan="6" class="vazio">Nenhuma máquina cadastrada.</td></tr>
    {% endfor %}
    <tr class="sem-resultado"><td colspan="6" class="vazio">Nenhum resultado.</td></tr>
    </tbody>
    </table>
    </div>
    </section>
</div>

<!-- ================= 3. PCM ================= -->
<div id="secao-pcm" class="painel-secao">
    <section class="resumo">
        <div class="resumo-card"><span class="icone">📅</span><span>Planos Ativos</span><strong>{{total_planos}}</strong></div>
        <div class="resumo-card"><span class="icone">✅</span><span>Em Dia</span><strong style="color:#4ade80">{{agendados}}</strong></div>
        <div class="resumo-card"><span class="icone">🚨</span><span>Atrasados</span><strong style="color:#f87171">{{atrasados}}</strong></div>
        <div class="resumo-card"><span class="icone">📊</span><span>Arquivos Excel</span><strong style="color:#60a5fa">{{arquivos_pcm|length}}</strong></div>
    </section>

    <details class="card"><summary>➕ Cadastrar Plano Preventivo</summary><div class="conteudo">
    <form action="{{url_for('cadastrar_plano_pcm')}}" method="POST"><div class="form-grid tres">
        <div class="campo"><label>Máquina / Equipamento *</label><input name="maquina" required maxlength="150" placeholder="Ex: Extrusora Troester"></div>
        <div class="campo"><label>Tag / Código</label><input name="codigo_maquina" maxlength="50" placeholder="Ex: EXT-01" style="text-transform:uppercase"></div>
        <div class="campo"><label>Setor</label><input name="setor" maxlength="100" placeholder="Ex: Extrusão"></div>
        <div class="campo"><label>Periodicidade *</label>
            <select name="frequencia" required>
                <option value="Semanal">Semanal (7d)</option>
                <option value="Quinzenal">Quinzenal (15d)</option>
                <option value="Mensal" selected>Mensal (30d)</option>
                <option value="Bimestral">Bimestral (2m)</option>
                <option value="Trimestral">Trimestral (3m)</option>
                <option value="Semestral">Semestral (6m)</option>
                <option value="Anual">Anual (12m)</option>
            </select>
        </div>
        <div class="campo"><label>Responsável</label><input name="responsavel" maxlength="100" placeholder="Ex: Mecânica"></div>
        <div class="campo"><label>Data Programada *</label><input type="date" name="proxima_execucao" required></div>
        <div class="campo completo"><label>Roteiro / Check-list Preventivo *</label><textarea name="atividade" data-max="2000" required placeholder="Pontos de inspeção, lubrificação..."></textarea></div>
        <div class="campo completo"><button class="salvar" style="width:100%" type="submit">+ Salvar Plano Preventivo</button></div>
    </div></form></div></details>

    <section class="card">
        <h2>📥 Arquivos Excel de PCM</h2>
        <form action="{{url_for('importar_excel_pcm')}}" method="POST" enctype="multipart/form-data">
            <div class="form-grid" style="align-items:end">
                <div class="campo">
                    <label>Selecionar Arquivo Excel (.xlsx / .xls) *</label>
                    <input type="file" name="arquivo_excel" accept=".xlsx,.xls" required>
                </div>
                <div class="campo">
                    <button type="submit" class="salvar" style="width:100%">⬆ Guardar Arquivo Excel</button>
                </div>
            </div>
        </form>

        {% if arquivos_pcm %}
        <div class="lista-arquivos-pcm">
            {% for aq in arquivos_pcm %}
            <div class="item-arquivo-pcm">
                <div class="item-arquivo-info">
                    <strong>📊 {{aq.nome_original}}</strong>
                    <small>{{aq.data_envio}} • {{aq.tamanho_kb}} KB</small>
                </div>
                <div class="item-arquivo-acoes">
                    <a href="{{url_for('abrir_arquivo_pcm', id=aq.id)}}" class="botao excel">📊 Abrir Excel</a>
                    <form action="{{url_for('excluir_arquivo_pcm', id=aq.id)}}" method="POST" data-perigo="1" data-confirm="Remover o arquivo {{aq.nome_original}}?" data-confirm-btn="Remover">
                        <button type="submit" class="excluir">✕</button>
                    </form>
                </div>
            </div>
            {% endfor %}
        </div>
        {% else %}
        <p class="muted" style="margin-top:12px;margin-bottom:0">Nenhum arquivo Excel guardado no PCM.</p>
        {% endif %}
    </section>

    <section class="card">
        <div class="acoes-topo">
            <div class="busca"><div class="campo-busca"><input class="q-busca" placeholder="Pesquisar preventiva..."></div></div>
            <a class="botao excel" href="{{url_for('excel_pcm')}}">⬇ Exportar Excel (PCM)</a>
        </div>
    </section>

    <section class="card">
    <h2>Cronograma de Manutenção Preventiva</h2>
    <div class="tabela-wrapper">
    <table id="tabelaPCM" class="tabela-dinamica detalhavel">
    <thead><tr><th data-tipo="texto">Máquina</th><th data-tipo="texto">Tag</th><th data-tipo="texto">Setor</th><th data-tipo="texto">Frequência</th><th>Próxima Exec</th><th data-tipo="texto">Resp.</th><th>Status</th><th>Ações</th></tr></thead>
    <tbody>
    {% for p in planos %}
    <tr class="linha">
        <td data-label="Máquina"><strong>{{p.maquina}}</strong></td>
        <td data-label="Tag">{{p.codigo_maquina or '-'}}</td>
        <td data-label="Setor">{{p.setor or '-'}}</td>
        <td data-label="Frequência"><span class="tipo-badge tipo-preventiva">{{p.frequencia}}</span></td>
        <td data-label="Próxima Exec"><strong class="{{ 'vencido' if p.status_calc=='Atrasado' else '' }}">{{p.proxima_execucao}}</strong></td>
        <td data-label="Responsável">{{p.responsavel or '-'}}</td>
        <td data-label="Status"><span class="status {{ 'estoque-baixo' if p.status_calc=='Atrasado' else 'estoque-ok' }}">{{p.status_calc}}</span></td>
        <td data-label="Ações">
            <div class="acoes">
                <form action="{{url_for('gerar_os_pcm', id=p.id)}}" method="POST" data-confirm="Gerar O.S. aberta para esta preventiva?" data-confirm-btn="Gerar O.S.">
                    <button class="imprimir" title="Abrir O.S.">📋 Abrir O.S.</button>
                </form>
                <form action="{{url_for('concluir_plano_pcm', id=p.id)}}" method="POST" data-confirm="Concluir preventivo e reagendar?" data-confirm-btn="Concluir">
                    <button class="finalizar" title="Concluir ciclo">✔ Concluir</button>
                </form>
                <button type="button" class="editar-peca btn-editar-pcm" data-url="{{url_for('editar_plano_pcm', id=p.id)}}" data-maquina="{{p.maquina}}" data-codigo="{{p.codigo_maquina}}" data-setor="{{p.setor}}" data-frequencia="{{p.frequencia}}" data-responsavel="{{p.responsavel}}" data-proxima="{{p.proxima_execucao}}" data-atividade="{{p.atividade}}">✏️ Editar</button>
                <form action="{{url_for('excluir_plano_pcm', id=p.id)}}" method="POST" data-perigo="1" data-confirm="Excluir plano?" data-confirm-btn="Excluir"><button class="excluir">✕ Excluir</button></form>
            </div>
        </td>
    </tr>
    <tr class="detalhe"><td colspan="8"><strong style="color:var(--or2)">Roteiro:</strong><p>{{p.atividade}}</p></td></tr>
    {% else %}
    <tr><td colspan="8" class="vazio">Nenhum plano preventivo.</td></tr>
    {% endfor %}
    <tr class="sem-resultado"><td colspan="8" class="vazio">Nenhum resultado.</td></tr>
    </tbody>
    </table>
    </div>
    </section>
</div>

<!-- ================= 4. COMPRAS (S.C.) ================= -->
<div id="secao-compras" class="painel-secao">
    <section class="resumo">
        <div class="resumo-card" data-atalho='[data-filtro="tipo"][data-valor=""]'><span class="icone">🛒</span><span>Total de S.C.</span><strong>{{total_sc}}</strong></div>
        <div class="resumo-card" data-atalho='[data-filtro="tipo"][data-valor="S.C. Emergencial"]'><span class="icone">🚨</span><span>S.C. Emergencial</span><strong style="color:#f87171">{{emergenciais}}</strong></div>
        <div class="resumo-card" data-atalho='[data-filtro="tipo"][data-valor="S.C. Programada"]'><span class="icone">📅</span><span>S.C. Programada</span><strong style="color:#60a5fa">{{programadas}}</strong></div>
        <div class="resumo-card" data-atalho='[data-filtro="tipo"][data-valor="S.C. Reposição / Estoque"]'><span class="icone">📦</span><span>S.C. Reposição</span><strong style="color:#c084fc">{{reposicao}}</strong></div>
    </section>

    <details class="card" open><summary>➕ Nova Solicitação de Compra (S.C.)</summary><div class="conteudo">
    <form action="{{url_for('cadastrar_compras')}}" method="POST"><div class="form-grid tres">
        <div class="campo"><label>Número / Código S.C. *</label><input name="numero_sc" value="{{proxima_sc}}" style="color:var(--or2);font-weight:700"></div>
        <div class="campo"><label>Data da Solicitação *</label><input type="date" name="data_solicitacao" value="{{hoje}}" required></div>
        <div class="campo"><label>Centro de Custo *</label>
            <select name="centro_custo" required>
                {% for cc in centros_custo_opcoes %}<option value="{{ cc }}" {% if cc == "20611 – VULCANIZAÇÃO" %}selected{% endif %}>{{ cc }}</option>{% endfor %}
            </select>
        </div>
        <div class="campo completo"><label>Classificação / Criticidade da S.C. *</label>
            <div class="seg">
                <input type="radio" id="scEmerg" name="tipo_sc" value="S.C. Emergencial" {% if not repor_item %}checked{% endif %}><label for="scEmerg">🚨 Emergencial</label>
                <input type="radio" id="scProg" name="tipo_sc" value="S.C. Programada"><label for="scProg">📅 Programada</label>
                <input type="radio" id="scRep" name="tipo_sc" value="S.C. Reposição / Estoque" {% if repor_item %}checked{% endif %}><label for="scRep">📦 Reposição</label>
            </div>
        </div>
        <div class="campo"><label>Solicitante *</label><input name="solicitante" required maxlength="100" value="Carlos"></div>
        <div class="campo"><label>Setor Solicitante</label><input name="setor_solicitante" maxlength="100" value="Manutenção"></div>
        <div class="campo"><label>Responsáveis Técnicos</label><input name="responsaveis_duvidas" maxlength="150" value="Carlos e Jonatas"></div>
        <div class="campo completo"><label>Objeto / Finalidade Geral *</label><input name="objeto_solicitacao" required maxlength="250" value="{{ 'Reposição de Estoque Mínimo — ' + repor_item.nome if repor_item else '' }}"></div>
        <div class="campo completo">
            <label style="color:var(--or2);font-size:14px;margin-bottom:8px">🛒 Itens da Solicitação</label>
            <div class="bloco-itens" id="containerItensSC">
                <div class="linha-item-sc">
                    <input type="text" name="descritivo_material[]" required value="{{ (repor_item.codigo + ' - ' + repor_item.nome) if repor_item else '' }}" placeholder="Descritivo Técnico *">
                    <div class="linha-item-sc-baixo">
                        <input type="number" name="quantidade[]" value="{{ repor_item.qtd if repor_item else 1 }}" min="1" required placeholder="Qtd">
                        <select name="unidade[]"><option value="Un">Un (Un)</option><option value="Pç">Peça (Pç)</option><option value="M">Metros (M)</option><option value="Kg">Kg (Kg)</option><option value="L">Litros (L)</option></select>
                        <input type="text" name="link_referencia[]" placeholder="Link / Referência">
                    </div>
                </div>
            </div>
            <button type="button" class="botao limpar" id="btnAddItemSC" style="margin-top:8px;width:100%">+ Adicionar Outro Item</button>
        </div>
        <div class="campo completo"><button class="salvar" style="width:100%" type="submit">+ Registrar Solicitação de Compra</button></div>
    </div></form></div></details>

    <section class="card">
        <div class="acoes-topo">
            <div class="busca"><div class="campo-busca"><input class="q-busca" placeholder="Pesquisar S.C..."></div></div>
            <a class="botao excel" href="{{url_for('excel_compras')}}">⬇ Exportar Excel (S.C.)</a>
        </div>
        <div class="filtros">
            <span class="rotulo">Criticidade</span>
            <button type="button" class="chip ativo" data-filtro="tipo" data-valor="">Todas</button>
            <button type="button" class="chip" data-filtro="tipo" data-valor="S.C. Emergencial">🚨 Emergencial</button>
            <button type="button" class="chip" data-filtro="tipo" data-valor="S.C. Programada">📅 Programada</button>
            <button type="button" class="chip" data-filtro="tipo" data-valor="S.C. Reposição / Estoque">📦 Reposição</button>
            <span class="contagem" id="contador-tabelaCompras"></span>
        </div>
    </section>

    <section class="card">
    <h2>Solicitações de Compras Registradas</h2>
    <div class="tabela-wrapper">
    <table id="tabelaCompras" class="tabela-dinamica detalhavel">
    <thead><tr><th>Código S.C.</th><th>Criticidade</th><th>Solicitante</th><th>Centro Custo</th><th>Objeto</th><th>Qtd</th><th>Data</th><th>Status</th><th>Ações</th></tr></thead>
    <tbody>
    {% for sc in compras %}
    <tr class="linha" data-tipo="{{sc.tipo_sc}}">
        <td data-label="Código S.C."><strong>{{sc.numero_sc}}</strong></td>
        <td data-label="Criticidade">
            {% if sc.tipo_sc == 'S.C. Emergencial' %}<span class="sc-badge-emergencial">🚨 Emergencial</span>
            {% elif sc.tipo_sc == 'S.C. Programada' %}<span class="sc-badge-programada">📅 Programada</span>
            {% else %}<span class="sc-badge-reposicao">📦 Reposição</span>{% endif %}
        </td>
        <td data-label="Solicitante">{{sc.solicitante}}</td>
        <td data-label="Centro Custo">{{sc.centro_custo or '-'}}</td>
        <td data-label="Objeto"><strong>{{sc.objeto_solicitacao}}</strong> <span class="muted">({{sc.itens|length}} it)</span></td>
        <td data-label="Qtd Total">{{ sc.itens|sum(attribute='quantidade') }}</td>
        <td data-label="Data">{{sc.data_solicitacao}}</td>
        <td data-label="Status">
            <select class="select-status-sc" data-sc-id="{{sc.id}}" data-val="{{sc.status_aprovacao}}">
                <option value="Aguardando Aprovação" {% if sc.status_aprovacao == 'Aguardando Aprovação' %}selected{% endif %}>🟡 Aguardando</option>
                <option value="Em Cotação" {% if sc.status_aprovacao == 'Em Cotação' %}selected{% endif %}>🔵 Em Cotação</option>
                <option value="Aprovado" {% if sc.status_aprovacao == 'Aprovado' %}selected{% endif %}>🟢 Aprovado</option>
                <option value="Reprovado" {% if sc.status_aprovacao == 'Reprovado' %}selected{% endif %}>🔴 Reprovado</option>
                <option value="Finalizado / Comprado" {% if sc.status_aprovacao == 'Finalizado / Comprado' %}selected{% endif %}>🟣 Comprado</option>
            </select>
        </td>
        <td data-label="Ações">
            <div class="acoes">
                <a href="{{url_for('imprimir_sc', id=sc.id)}}" target="_blank" class="botao imprimir" title="Imprimir A4">🖨️ A4</a>
                <form action="{{url_for('excluir_compras', id=sc.id)}}" method="POST" data-perigo="1" data-confirm="Excluir S.C. {{sc.numero_sc}}?" data-confirm-btn="Excluir"><button class="excluir">✕ Excluir</button></form>
            </div>
        </td>
    </tr>
    <tr class="detalhe"><td colspan="9"><strong style="color:var(--or2)">Itens:</strong>
        <div class="tabela-wrapper" style="overflow-x:auto;margin-top:8px">
        <table style="background:#14171f;min-width:560px"><thead><tr><th>#</th><th>Descritivo</th><th>Qtd</th><th>Unid</th><th>Ref / Link</th></tr></thead><tbody>
        {% for it in sc.itens %}
        {% set ref = (it.link_referencia or '')|trim %}
        <tr>
            <td>{{loop.index}}</td>
            <td><strong>{{it.descritivo}}</strong></td>
            <td>{{it.quantidade}}</td>
            <td>{{it.unidade}}</td>
            <td>
                {% if ref.startswith('http://') or ref.startswith('https://') %}
                    <a href="{{ref}}" target="_blank" rel="noopener noreferrer" class="botao imprimir" style="padding:5px 10px;font-size:11.5px" title="{{ref}}">🔗 Abrir Link</a>
                {% elif ref.startswith('www.') %}
                    <a href="https://{{ref}}" target="_blank" rel="noopener noreferrer" class="botao imprimir" style="padding:5px 10px;font-size:11.5px" title="https://{{ref}}">🔗 Abrir Link</a>
                {% else %}
                    {{ref or '-'}}
                {% endif %}
            </td>
        </tr>
        {% endfor %}
        </tbody></table>
        </div>
    </td></tr>
    {% else %}
    <tr><td colspan="9" class="vazio">Nenhuma S.C. cadastrada.</td></tr>
    {% endfor %}
    <tr class="sem-resultado"><td colspan="9" class="vazio">Nenhum resultado.</td></tr>
    </tbody>
    </table>
    </div>
    </section>
</div>

<!-- ================= 5. INDICADORES, MTTR & EQUIPE ================= -->
<div id="secao-indicadores" class="painel-secao">
    <section class="resumo">
        <div class="resumo-card"><span class="icone">🛡️</span><span>Índice Preventivo</span><strong style="color:#60a5fa">{{pct_geral_prev}}%</strong><small>{{total_preventiva}} prev vs {{total_corretiva}} corr</small></div>
        <div class="resumo-card"><span class="icone">⏱️</span><span>MTTR (Tempo Reparo)</span><strong style="color:#fb923c" id="valorMTTR">{{mttr_str}}</strong><small id="subMTTR">Média geral corretivas</small></div>
        <div class="resumo-card">
            <span class="icone">🔄</span>
            <span>MTBF por Máquina</span>
            <strong style="color:#4ade80" id="valorMTBF">{{mtbf_str}}</strong>
            <select id="selectMaquinaMTBF" style="margin-top:8px;padding:5px 9px;font-size:13px;border-radius:8px;min-height:38px">
                <option value="geral">Todas as máquinas (Média)</option>
                {% for m in mtbf_maquinas %}
                <option value="{{m.chave}}">{{m.nome}} ({{m.corretivas}} falhas)</option>
                {% endfor %}
            </select>
            <small id="subMTBF">Média entre falhas</small>
        </div>
        <div class="resumo-card"><span class="icone">📈</span><span>O.S. Auditadas</span><strong>{{total_fin}}</strong><small>{{total_meses}} mês(es)</small></div>
    </section>

    <div class="grid-duplo">
        <section class="card" style="margin-bottom:0">
            <h2>Proporção Preventiva vs. Corretiva</h2>
            <div style="display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap">
                <label style="margin:0;color:var(--mt)">Período:</label>
                <select id="selectMes" style="width:auto;padding:6px 12px;border-radius:8px">
                    <option value="geral">📅 Visão Geral (Acumulado)</option>
                    {% for m in dados_tabela %}<option value="{{m.chave}}">{{m.mes_nome}}</option>{% endfor %}
                </select>
            </div>
            <div style="position:relative;width:100%;max-width:380px;height:280px;margin:0 auto"><canvas id="graficoPizza"></canvas></div>
        </section>

        <section class="card" style="margin-bottom:0">
            <h2>🏭 Confiabilidade por Máquina</h2>
            {% if mtbf_maquinas %}
            <div class="tabela-wrapper" style="margin-top:10px;overflow-x:auto">
                <table style="min-width:540px">
                    <thead><tr><th>Máquina</th><th>Corretivas</th><th>MTBF</th><th>MTTR</th><th>Horas Totais</th></tr></thead>
                    <tbody>
                    {% for m in mtbf_maquinas %}
                    <tr>
                        <td><strong>{{m.nome}}</strong></td>
                        <td><span class="tipo-badge tipo-corretiva">{{m.corretivas}}</span></td>
                        <td><strong style="color:#4ade80">{{m.mtbf_str}}</strong></td>
                        <td><strong style="color:#fb923c">{{m.mttr_str}}</strong></td>
                        <td>{{m.horas_str}}</td>
                    </tr>
                    {% endfor %}
                    </tbody>
                </table>
            </div>
            {% else %}
            <p class="muted" style="margin-top:20px">Nenhuma O.S. corretiva finalizada para cálculo por máquina.</p>
            {% endif %}
        </section>
    </div>

    <section class="card">
        <h2>👷 Eficiência Individual dos Funcionários</h2>
        <div style="display:flex;gap:12px;align-items:center;margin-bottom:18px;flex-wrap:wrap">
            <form method="GET" style="display:flex;gap:8px;align-items:center;margin:0">
                <input type="hidden" name="tab" value="indicadores">
                <label for="mes" style="margin:0;color:var(--mt)">Período:</label>
                <select name="mes" id="mes" style="width:auto" onchange="this.form.submit()">
                    <option value="geral" {{ 'selected' if mes_sel=='geral' else '' }}>📅 Todo o período</option>
                    {% for m in meses %}<option value="{{m.chave}}" {{ 'selected' if m.chave==mes_sel else '' }}>{{m.nome}}</option>{% endfor %}
                </select>
            </form>
            <select id="filtroFuncaoEquipe" style="width:auto">
                <option value="">Todas as funções ({{funcionarios_indiv|length}})</option>
                <option value="mecanica">🔧 Mecânica & Auxiliar</option>
                <option value="eletrica">⚡ Elétrica</option>
                <option value="predial">🏢 Predial</option>
            </select>
        </div>
        <div class="grid-triplo" id="gridFuncionarios">
        {% for f in funcionarios_indiv %}
            {% set cat_func = 'eletrica' if f.funcao == 'Eletricista' else ('predial' if f.funcao == 'Predial' else 'mecanica') %}
            {% set badge_cls = 'tipo-preventiva' if f.funcao == 'Eletricista' else ('tipo-predial' if f.funcao == 'Predial' else 'tipo-corretiva') %}
            <div class="card-funcionario" data-cat="{{cat_func}}" style="background:#0c0e12;border:1px solid var(--bd);border-radius:12px;padding:16px">
                <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:6px">
                    <h3 style="margin:0;color:var(--or2);font-size:16px">{{f.icone}} {{f.nome}}</h3>
                    <span class="tipo-badge {{badge_cls}}">{{f.funcao}}</span>
                </div>
                <p style="margin:3px 0;font-size:12px;color:var(--mt)"><strong>Equipe:</strong> {{f.equipe}} • <strong>Turno:</strong> {{f.turno}}</p>
                <p style="margin:6px 0;font-size:13px">Apontado: <strong style="color:var(--or2)">{{ f.tempo_str }}</strong> ({{ f.qtd_os }} O.S. • {{ f.dias }} dia(s))</p>
                <p style="margin:6px 0;font-size:13px">Eficiência Individual: <strong style="color:#4ade80">{{ f.eficiencia }}%</strong></p>
                <div class="progresso"><span style="width:{{ f.eficiencia_clamp }}%"></span></div>
            </div>
        {% endfor %}
        </div>
    </section>
</div>

<!-- ================= 6. BACKUP & AUDITORIA ================= -->
<div id="secao-governanca" class="painel-secao">
    <section class="card">
        <h2>💾 Cópia de Segurança do Banco de Dados</h2>
        <div class="resumo" style="margin:16px 0">
            <div class="resumo-card"><span class="icone">🗄️</span><span>Motor Ativo</span><strong style="font-size:20px">{{ 'PostgreSQL' if is_postgres else 'SQLite Local' }}</strong><small>{{ 'Nuvem Neon' if is_postgres else (tamanho_db ~ ' KB') }}</small></div>
            <div class="resumo-card"><span class="icone">📋</span><span>O.S. Salvas</span><strong>{{resumo_banco.ordens}}</strong><small>Histórico geral</small></div>
            <div class="resumo-card"><span class="icone">🔩</span><span>Itens & Ativos</span><strong>{{resumo_banco.estoque + resumo_banco.maquinas}}</strong><small>{{resumo_banco.estoque}} SKUs • {{resumo_banco.maquinas}} Máq.</small></div>
            <div class="resumo-card"><span class="icone">📅</span><span>PCM & Compras</span><strong>{{resumo_banco.planos + resumo_banco.compras}}</strong><small>{{resumo_banco.planos}} Planos • {{resumo_banco.compras}} S.C.</small></div>
        </div>

        <div style="display:flex;gap:12px;flex-wrap:wrap;margin-top:16px">
            <a href="{{url_for('exportar_backup_json')}}" class="botao salvar">📦 Exportar Banco Completo (.JSON)</a>
            <a href="{{url_for('exportar_backup_excel_completo')}}" class="botao excel">📊 Exportar Planilha Geral Multi-Abas (.XLSX)</a>
            {% if not is_postgres %}
            <a href="{{url_for('baixar_backup_db')}}" class="botao imprimir">💾 Baixar Arquivo Bruto SQLite (.DB)</a>
            {% endif %}
        </div>
    </section>

    <section class="card">
        <h2>🛡️ Trilha de Auditoria</h2>
        <div class="tabela-wrapper" style="overflow-x:auto">
            <table style="min-width:680px">
                <thead><tr><th>Data / Hora</th><th>Módulo</th><th>Ação</th><th>Detalhes</th></tr></thead>
                <tbody>
                {% for l in logs %}
                    <tr><td>{{l.data_hora}}</td><td><span class="tipo-badge tipo-preventiva">{{l.modulo}}</span></td><td><strong>{{l.acao}}</strong></td><td>{{l.detalhes}}</td></tr>
                {% else %}
                    <tr><td colspan="4" class="vazio">Nenhum evento registrado.</td></tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
    </section>
</div>

<footer class="rodape">Central de Gestão de Manutenção — JAMAICA</footer></div>

<nav class="mobile-bottom-nav" aria-label="Navegação Rápida Mobile">
    <a href="{{url_for('lobby')}}"><span class="nav-ico">🏠</span>Lobby</a>
    <a href="{{url_for('dashboard')}}"><span class="nav-ico">📋</span>O.S.</a>
    <a href="{{url_for('central_gestao')}}" class="ativo"><span class="nav-ico">🔐</span>Gestão</a>
    <a href="{{url_for('dashboard')}}#cardNovaOS"><span class="nav-ico">➕</span>Nova O.S.</a>
</nav>

<div class="overlay" id="modalMov"><div class="modal"><h3 id="movTitulo"></h3><p class="muted" id="movSub"></p><form id="formMov" method="POST"><div class="seg"><input type="radio" id="tipoRet" name="tipo" value="Retirada" checked><label for="tipoRet">▼ Retirar</label><input type="radio" id="tipoEnt" name="tipo" value="Entrada"><label for="tipoEnt">▲ Adicionar</label></div><div class="campo"><label>Quantidade</label><div class="stepper"><button type="button" data-step="-1">−</button><input type="number" name="qtd" id="movQtd" min="1" value="1" required><button type="button" data-step="1">+</button></div></div><p class="saldo" id="boxSaldo">Saldo após a operação: <strong id="movSaldo">0</strong></p><div class="campo"><label id="movLblSol">Solicitante / O.S. *</label><input type="text" name="solicitante" id="movSol" maxlength="100" autocomplete="off"></div><div class="modal-acoes"><button type="button" class="limpar" data-fechar>Cancelar</button><button type="submit" class="salvar">Confirmar</button></div></form></div></div>

<div class="overlay" id="modalEditarPeca">
    <div class="modal">
        <h3>✏️ Editar Cadastro da Peça</h3>
        <form id="formEditarPeca" method="POST">
            <div class="campo"><label>Código / SKU *</label><input type="text" name="codigo" id="editarCodigoPeca" maxlength="50" required style="text-transform:uppercase"></div>
            <div class="campo"><label>Nome da Peça *</label><input type="text" name="nome" id="editarNomePeca" maxlength="150" required></div>
            <div class="campo"><label>Categoria</label><input type="text" name="categoria" id="editarCategoriaPeca" maxlength="100"></div>
            <div class="campo"><label>Localização</label><input type="text" name="localizacao" id="editarLocalizacaoPeca" maxlength="100"></div>
            <div class="campo"><label>Estoque Mínimo *</label><input type="number" name="estoque_minimo" id="editarEstoqueMinimoPeca" min="1" required></div>
            <div class="modal-acoes"><button type="button" class="limpar" id="cancelarEdicaoPeca">Cancelar</button><button type="submit" class="salvar">Salvar</button></div>
        </form>
    </div>
</div>

<div class="overlay" id="modalEditarSpare">
    <div class="modal">
        <h3>✏️ Editar Máquina / Equipamento (Spare Parts)</h3>
        <form id="formEditarSpare" method="POST">
            <div class="campo"><label>Máquina / Equipamento *</label><input type="text" name="maquina" id="spareMaquinaInput" maxlength="150" required></div>
            <div class="campo"><label>Modelo / Fabricante *</label><input type="text" name="modelo" id="spareModeloInput" maxlength="100" required></div>
            <div class="campo"><label>Setor</label><input type="text" name="setor" id="spareSetorInput" maxlength="100"></div>
            <div class="campo"><label>Notas Técnicas / Peças Críticas Recomendadas</label><textarea name="pecas_necessarias" id="spareNotasInput" rows="3"></textarea></div>
            <div class="modal-acoes">
                <button type="button" class="limpar" id="cancelarEdicaoSpare">Cancelar</button>
                <button type="submit" class="salvar">Salvar Alterações</button>
            </div>
        </form>
    </div>
</div>

<div class="overlay" id="modalEditarPCM">
    <div class="modal">
        <h3>✏️ Editar Planejamento Preventivo</h3>
        <form id="formEditarPCM" method="POST">
            <div class="form-grid tres">
                <div class="campo"><label>Máquina *</label><input type="text" name="maquina" id="pcmMaquina" required></div>
                <div class="campo"><label>Tag</label><input type="text" name="codigo_maquina" id="pcmCodigo" style="text-transform:uppercase"></div>
                <div class="campo"><label>Setor</label><input type="text" name="setor" id="pcmSetor"></div>
                <div class="campo"><label>Frequência *</label>
                    <select name="frequencia" id="pcmFrequencia" required><option value="Semanal">Semanal</option><option value="Quinzenal">Quinzenal</option><option value="Mensal">Mensal</option><option value="Bimestral">Bimestral</option><option value="Trimestral">Trimestral</option><option value="Semestral">Semestral</option><option value="Anual">Anual</option></select>
                </div>
                <div class="campo"><label>Responsável</label><input type="text" name="responsavel" id="pcmResponsavel"></div>
                <div class="campo"><label>Próxima Execução *</label><input type="date" name="proxima_execucao" id="pcmProxima" required></div>
                <div class="campo completo"><label>Roteiro Preventivo *</label><textarea name="atividade" id="pcmAtividade" rows="5" required></textarea></div>
            </div>
            <div class="modal-acoes"><button type="button" class="limpar" id="fecharModalPCM">Cancelar</button><button type="submit" class="salvar">Salvar</button></div>
        </form>
    </div>
</div>

<script>
window.__FLASH__={{flash_msgs|tojson}};
window.__ULTIMO_OS_ID__={{ultimo_os_id|default(0)}};
const dadosTabela = {{ dados_tabela|tojson }};
const totalGeralPrev = {{ total_preventiva }};
const totalGeralCorr = {{ total_corretiva }};
const mtbfMaquinasDados = {{ mtbf_maquinas|tojson }};
const mtbfGeralTexto = {{ mtbf_str|tojson }};
const mttrGeralTexto = {{ mttr_str|tojson }};

document.addEventListener("DOMContentLoaded", function() {
    const ctx = document.getElementById('graficoPizza').getContext('2d');
    const chartInstance = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Preventiva', 'Corretiva'],
            datasets: [{
                data: [totalGeralPrev, totalGeralCorr],
                backgroundColor: ['rgba(59, 130, 246, 0.85)', 'rgba(249, 115, 22, 0.85)'],
                borderColor: ['#3b82f6', '#f97316'],
                borderWidth: 2,
                hoverOffset: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: '#f3f4f6', font: { size: 12 }, padding: 12 } } },
            cutout: '58%'
        }
    });

    document.getElementById('selectMes')?.addEventListener('change', function() {
        const val = this.value;
        if(val === 'geral') { chartInstance.data.datasets[0].data = [totalGeralPrev, totalGeralCorr]; }
        else {
            const encontrado = dadosTabela.find(item => item.chave === val);
            chartInstance.data.datasets[0].data = encontrado ? [encontrado.prev_qtd, encontrado.corr_qtd] : [0, 0];
        }
        chartInstance.update();
    });

    document.getElementById('filtroFuncaoEquipe')?.addEventListener('change', function() {
        const cat = this.value;
        document.querySelectorAll('#gridFuncionarios .card-funcionario').forEach(card => {
            card.style.display = (!cat || card.dataset.cat === cat) ? '' : 'none';
        });
    });

    document.getElementById('selectMaquinaMTBF')?.addEventListener('change', function() {
        const val = this.value;
        const elMtbf = document.getElementById('valorMTBF');
        const elSubMtbf = document.getElementById('subMTBF');
        const elMttr = document.getElementById('valorMTTR');
        const elSubMttr = document.getElementById('subMTTR');

        if (val === 'geral') {
            elMtbf.textContent = mtbfGeralTexto;
            elSubMtbf.textContent = 'Média entre máquinas com 2+ falhas';
            elMttr.textContent = mttrGeralTexto;
            elSubMttr.textContent = 'Média geral corretivas';
        } else {
            const item = mtbfMaquinasDados.find(m => m.chave === val);
            if (item) {
                elMtbf.textContent = item.mtbf_str;
                elSubMtbf.textContent = item.corretivas >= 2
                    ? `${item.nome} (${item.corretivas} falhas registradas)`
                    : `${item.nome} (requer mín. 2 falhas)`;
                elMttr.textContent = item.mttr_str;
                elSubMttr.textContent = `Reparo médio: ${item.nome}`;
            }
        }
    });

    const modalEditarPCM = document.getElementById("modalEditarPCM");
    const formEditarPCM = document.getElementById("formEditarPCM");
    document.querySelectorAll(".btn-editar-pcm").forEach(btn => {
        btn.addEventListener("click", function() {
            formEditarPCM.action = this.dataset.url;
            document.getElementById("pcmMaquina").value = this.dataset.maquina || "";
            document.getElementById("pcmCodigo").value = this.dataset.codigo || "";
            document.getElementById("pcmSetor").value = this.dataset.setor || "";
            document.getElementById("pcmFrequencia").value = this.dataset.frequencia || "Mensal";
            document.getElementById("pcmResponsavel").value = this.dataset.responsavel || "";
            let dt = this.dataset.proxima || "";
            if (dt.includes("/")) { let p = dt.split("/"); if (p.length === 3) dt = `${p[2]}-${p[1]}-${p[0]}`; }
            document.getElementById("pcmProxima").value = dt;
            document.getElementById("pcmAtividade").value = this.dataset.atividade || "";
            modalEditarPCM.classList.add("aberto");
        });
    });
    document.getElementById("fecharModalPCM")?.addEventListener("click", () => modalEditarPCM.classList.remove("aberto"));

    const containerItens = document.getElementById("containerItensSC");
    const btnAddItem = document.getElementById("btnAddItemSC");
    if (btnAddItem) {
        btnAddItem.addEventListener("click", function() {
            const novaLinha = document.createElement("div");
            novaLinha.className = "linha-item-sc";
            novaLinha.innerHTML = `
                <input type="text" name="descritivo_material[]" required placeholder="Descritivo Técnico *">
                <div class="linha-item-sc-baixo">
                    <input type="number" name="quantidade[]" value="1" min="1" required placeholder="Qtd">
                    <select name="unidade[]"><option value="Un">Un (Un)</option><option value="Pç">Peça (Pç)</option><option value="M">Metros (M)</option><option value="Kg">Kg (Kg)</option><option value="L">Litros (L)</option></select>
                    <input type="text" name="link_referencia[]" placeholder="Link / Referência">
                    <button type="button" class="excluir btn-rm-item" style="padding:8px" title="Remover">✕</button>
                </div>
            `;
            containerItens.appendChild(novaLinha);
            novaLinha.querySelector(".btn-rm-item").addEventListener("click", () => novaLinha.remove());
        });
    }

    document.querySelectorAll(".select-status-sc").forEach(sel => {
        sel.addEventListener("change", function() {
            const scId = this.dataset.scId;
            const novoStatus = this.value;
            this.dataset.val = novoStatus;
            fetch(`/compras/atualizar-status/${scId}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'X-CSRF-Token': window.__CSRF__ || '' },
                body: new URLSearchParams({ 'status_aprovacao': novoStatus, 'csrf_token': window.__CSRF__ || '' })
            }).then(res => res.json()).then(data => {
                if(data.ok) { window.toast(`Status atualizado para: ${novoStatus}`, 'ok'); }
                else { window.toast(data.erro || 'Erro ao atualizar.', 'erro'); }
            }).catch(() => window.toast('Erro de rede.', 'erro'));
        });
    });

    const urlParams = new URLSearchParams(window.location.search);
    const tabParam = urlParams.get('tab');
    if(tabParam && document.querySelector('[data-alvo="secao-'+tabParam+'"]')) {
        document.querySelector('[data-alvo="secao-'+tabParam+'"]').click();
    }
});
</script>
<script>{{js|safe}}</script></body></html>
"""

HTML_IMPRESSAO_OS = r"""
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>O.S. {{os.numero_os}} — Impressão</title>
<style>
@page { size: A4 portrait; margin: 10mm; }
body { font-family: Arial, sans-serif; color: #111; margin: 0; padding: 0; font-size: 11.5px; line-height: 1.35; }
.a4-page { width: 100%; max-width: 190mm; margin: 0 auto; box-sizing: border-box; }
.box { border: 2px solid #111; padding: 14px 16px; margin-bottom: 12px; background: #fff; }
.topo { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #111; padding-bottom: 10px; margin-bottom: 12px; }
.topo h1 { margin: 0; font-size: 16px; font-weight: bold; }
.topo small { font-size: 11px; color: #444; }
.badge { border: 2px solid #111; padding: 6px 12px; font-weight: bold; font-size: 14px; background: #eee; }
table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
td, th { border: 1px solid #333; padding: 6px 8px; text-align: left; vertical-align: top; }
th { background: #f0f0f0; font-size: 11px; font-weight: bold; width: 22%; }
.area-grande { min-height: 75px; white-space: pre-wrap; }
.secao-titulo { font-weight: bold; font-size: 12px; background: #e5e7eb; padding: 5px 8px; border: 1px solid #333; margin-bottom: -1px; }
.evidencia-print-grid { border: 1px solid #333; padding: 8px; margin-bottom: 10px; background: #fafafa; display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; }
.evidencia-print-item { text-align: center; border: 1px solid #ccc; padding: 4px; background: #fff; page-break-inside: avoid; }
.evidencia-print-item img { max-width: 100%; max-height: 175px; height: auto; }
.check-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 6px; padding: 8px; border: 1px solid #333; margin-bottom: 10px; background: #fafafa; font-size: 11px; }
.check-item { display: flex; align-items: center; gap: 6px; }
.assinaturas { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; margin-top: 22px; text-align: center; }
.assinaturas div { border-top: 1px solid #111; padding-top: 6px; font-size: 11px; font-weight: bold; }
@media print { .no-print { display: none !important; } body { padding: 0; } }
</style></head><body>
<div class="no-print" style="max-width:190mm;margin:0 auto 14mm;text-align:right">
    <button onclick="window.print()" style="padding:10px 18px;background:#ea580c;color:#fff;border:0;border-radius:6px;font-weight:bold;cursor:pointer">🖨️ Imprimir Folha A4 (O.S.)</button>
</div>
<div class="a4-page">
    <div class="box">
        <div class="topo">
            <div><h1>ORDEM DE SERVIÇO DE MANUTENÇÃO INDUSTRIAL</h1><small>JAMAICA — Sistema Integrado de Confiabilidade</small></div>
            <div class="badge">O.S. Nº {{os.numero_os}}</div>
        </div>
        <div class="secao-titulo">1. DADOS DA SOLICITAÇÃO E IDENTIFICAÇÃO DO ATIVO</div>
        <table>
            <tr><th>Data Abertura</th><td>{{os.abertura}}</td><th>Status Atual</th><td><strong>{{os.status}}</strong></td></tr>
            <tr><th>Solicitante</th><td>{{os.solicitante}}</td><th>Prioridade</th><td><strong style="color: {% if os.prioridade=='Crítica' %}#dc2626{% elif os.prioridade=='Alta' %}#d97706{% else %}#16a34a{% endif %}">{{os.prioridade}}</strong></td></tr>
            <tr><th>Máquina / Equipamento</th><td><strong>{{os.setor or '-'}}</strong></td><th>Setor</th><td>{{os.responsavel or '-'}}</td></tr>
            <tr><th>Causa / Sintoma Relatado</th><td colspan="3"><strong>{{os.titulo or '-'}}</strong></td></tr>
            {% if os.motivo_pausa and os.status=='Aguardando Peça' %}
            <tr><th>Motivo da Espera (S.C.)</th><td colspan="3" style="background:#fef3c7"><strong>⏳ {{os.motivo_pausa}}</strong></td></tr>
            {% endif %}
        </table>
        <div class="secao-titulo">2. APONTAMENTO TÉCNICO E EXECUÇÃO</div>
        <table>
            <tr><th>Tipo de Manutenção</th><td>{{os.tipo_manutencao or '☐ Preventiva      ☐ Corretiva'}}</td><th>Data / Hora Conclusão</th><td>{{os.finalizacao or '____/____/20____ __:__'}}</td></tr>
            <tr><th>Equipe Executante</th><td>{{os.funcionarios or '___________________________'}}</td><th>Tempo Apontado</th><td>{{fmt_hhmm(os.horas_trabalhadas) if os.horas_trabalhadas else '____h ____min'}}</td></tr>
            <tr><th>Laudo Técnico / Serviço Efetuado</th><td colspan="3" class="area-grande">{{os.observacao_tecnica or ''}}</td></tr>
        </table>
        {% set lista_ev = os.evidencias_lista %}
        {% if lista_ev %}
        <div class="secao-titulo">EVIDÊNCIAS FOTOGRÁFICAS DO SERVIÇO EXECUTADO ({{lista_ev|length}})</div>
        <div class="evidencia-print-grid">
            {% for ev in lista_ev %}
            <div class="evidencia-print-item">
                <img src="{{ev.url}}" alt="Evidência {{loop.index}} da O.S. {{os.numero_os}}">
                <div style="font-size:9.5px;color:#444;margin-top:3px">{{loop.index}}. {{ev.nome}}</div>
            </div>
            {% endfor %}
        </div>
        {% endif %}
        <div class="secao-titulo">3. CHECK-LIST DE LIBERAÇÃO E SEGURANÇA (LOTO)</div>
        <div class="check-grid">
            <div class="check-item">☐ Bloqueio de Energia (LOTO)</div>
            <div class="check-item">☐ Limpeza da Área / Ferramentas</div>
            <div class="check-item">☐ Teste a Vazio / Operacional</div>
            <div class="check-item">☐ Verificação de Sensores</div>
            <div class="check-item">☐ Reaperto Geral Executado</div>
            <div class="check-item">☐ Liberação para Produção</div>
        </div>
        <div class="assinaturas">
            <div>Assinatura do Técnico / Executante</div>
            <div>Visto do Solicitante / Produção</div>
            <div>Aprovação / Visto PCM</div>
        </div>
    </div>
</div>
</body></html>
"""

HTML_IMPRESSAO_SC = r"""
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>{{sc.numero_sc}} — Impressão</title>
<style>
body{font-family:Arial,sans-serif;color:#111;margin:24px;font-size:13px}
.box{border:2px solid #111;padding:18px;max-width:820px;margin:auto}
.topo{display:flex;justify-content:space-between;align-items:center;border-bottom:2px solid #111;padding-bottom:12px;margin-bottom:16px}
.topo h1{margin:0;font-size:18px}
.badge{border:1px solid #111;padding:4px 10px;font-weight:bold;font-size:13px}
table{width:100%;border-collapse:collapse;margin-bottom:14px}
td,th{border:1px solid #444;padding:8px;text-align:left;vertical-align:top}
th{background:#eee}
.assinaturas{display:grid;grid-template-columns:1fr 1fr;gap:30px;margin-top:45px;text-align:center}
.assinaturas div{border-top:1px solid #111;padding-top:6px;font-size:12px}
.btn-link-print{display:inline-block;padding:4px 9px;background:#1d4ed8;color:#fff;text-decoration:none;border-radius:4px;font-size:11.5px;font-weight:bold}
@media print{.no-print{display:none}body{margin:0}}
</style></head><body>
<div class="no-print" style="max-width:820px;margin:0 auto 14mm;text-align:right">
    <button onclick="window.print()" style="padding:10px 18px;background:#ea580c;color:#fff;border:0;border-radius:6px;font-weight:bold;cursor:pointer">🖨️ Imprimir Solicitação de Compra</button>
</div>
<div class="box">
    <div class="topo">
        <div><h1>SOLICITAÇÃO DE COMPRA (S.C.) — JAMAICA</h1><small>Departamento de Manutenção Industrial</small></div>
        <div class="badge">{{sc.numero_sc}}</div>
    </div>
    <table>
        <tr><th style="width:20%">Data</th><td>{{sc.data_solicitacao}}</td><th style="width:20%">Criticidade</th><td><strong>{{sc.tipo_sc}}</strong></td></tr>
        <tr><th>Solicitante</th><td>{{sc.solicitante}} ({{sc.setor_solicitante}})</td><th>Status</th><td>{{sc.status_aprovacao}}</td></tr>
        <tr><th>Centro de Custo</th><td colspan="3">{{sc.centro_custo}}</td></tr>
        <tr><th>Objeto / Aplicação</th><td colspan="3"><strong>{{sc.objeto_solicitacao}}</strong></td></tr>
        <tr><th>Dúvidas Técnicas</th><td colspan="3">{{sc.responsaveis_duvidas or '-'}}</td></tr>
    </table>
    <table>
        <thead><tr><th style="width:6%">#</th><th>Descritivo Técnico do Material / Peça</th><th style="width:10%">Qtd</th><th style="width:10%">Unid.</th><th style="width:18%">Referência / Link</th></tr></thead>
        <tbody>
        {% for it in sc.itens %}
        {% set ref = (it.link_referencia or '')|trim %}
        <tr>
            <td>{{loop.index}}</td>
            <td>{{it.descritivo}}</td>
            <td>{{it.quantidade}}</td>
            <td>{{it.unidade}}</td>
            <td>
                {% if ref.startswith('http://') or ref.startswith('https://') %}
                    <a href="{{ref}}" target="_blank" rel="noopener noreferrer" class="btn-link-print" title="{{ref}}">🔗 Abrir Link</a>
                {% elif ref.startswith('www.') %}
                    <a href="https://{{ref}}" target="_blank" rel="noopener noreferrer" class="btn-link-print" title="https://{{ref}}">🔗 Abrir Link</a>
                {% else %}
                    {{ref or '-'}}
                {% endif %}
            </td>
        </tr>
        {% endfor %}
        </tbody>
    </table>
    <div class="assinaturas">
        <div>Solicitante / Coordenação de Manutenção</div>
        <div>Aprovação Gerência / Compras</div>
    </div>
</div>
</body></html>
"""

HTML_QRCODES = r"""
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Etiquetas QR Code — Ativos JAMAICA</title>
<script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"></script>
<style>
body{font-family:Arial,sans-serif;background:#f3f4f6;color:#111;margin:20px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:16px;max-width:1000px;margin:auto}
.etiqueta{background:#fff;border:2px solid #111;border-radius:10px;padding:16px;text-align:center;page-break-inside:avoid}
.etiqueta h3{margin:8px 0 4px;font-size:15px}
.etiqueta p{margin:0 0 10px;font-size:12px;color:#444}
.qr-box{display:flex;justify-content:center;margin:8px 0}
@media print{.no-print{display:none}body{background:#fff;margin:0}}
</style></head><body>
<div class="no-print" style="max-width:1000px;margin:0 auto 16px;display:flex;justify-content:space-between;align-items:center">
    <strong>📱 Etiquetas QR Code para Chão de Fábrica (Abertura Rápida de O.S.)</strong>
    <button onclick="window.print()" style="padding:10px 18px;background:#ea580c;color:#fff;border:0;border-radius:6px;font-weight:bold;cursor:pointer">🖨️ Imprimir Etiquetas</button>
</div>
<div class="grid">
{% for m in maquinas %}
    <div class="etiqueta">
        <div style="font-size:11px;font-weight:bold;color:#ea580c">MANUTENÇÃO JAMAICA</div>
        <h3>{{m.maquina}}</h3>
        <p>{{m.modelo}} {% if m.setor %}• {{m.setor}}{% endif %}</p>
        <div class="qr-box" id="qr-{{m.id}}" data-url="{{url_for('dashboard', abrir_maquina=m.maquina, _external=True)}}"></div>
        <small style="font-size:10px;color:#666">Aponte a câmera para abrir O.S.</small>
    </div>
{% else %}
    <p>Cadastre máquinas na aba Spare Parts para gerar etiquetas.</p>
{% endfor %}
</div>
<script>
document.querySelectorAll('.qr-box').forEach(el => {
    new QRCode(el, { text: el.dataset.url, width: 130, height: 130 });
});
</script>
</body></html>
"""


# ==============================================================================
# DECORATOR DE AUTENTICAÇÃO CENTRALIZADA
# ==============================================================================
def requer_login(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logado_sistema"):
            return redirect(url_for("login_sistema"))
        return f(*args, **kwargs)

    return decorated_function


@app.route("/login", methods=["GET", "POST"])
def login_sistema():
    if request.method == "POST":
        eh_ajax = (
            request.is_json
            or request.headers.get("X-Requested-With") == "XMLHttpRequest"
        )
        ip = request.remote_addr or "?"
        agora = time.time()
        falhas, desde = tentativas_login.get(ip, (0, agora))
        if agora - desde >= BLOQUEIO_LOGIN_SEG:
            falhas, desde = 0, agora
        if falhas >= MAX_TENTATIVAS_LOGIN:
            msg_bloq = "Muitas tentativas incorretas. Aguarde alguns minutos."
            if eh_ajax:
                return jsonify({"ok": False, "erro": msg_bloq}), 429
            flash(msg_bloq)
            return renderizar(
                HTML_LOGIN_SISTEMA,
                ultimo_os_id=(
                    db.session.query(db.func.max(OrdemServico.id)).scalar() or 0
                ),
            )

        payload = request.get_json(silent=True) if request.is_json else {}
        usuario = (
            (payload.get("usuario") if isinstance(payload, dict) else None)
            or request.form.get("usuario", "")
        ).strip()
        senha = (
            (payload.get("senha") if isinstance(payload, dict) else None)
            or request.form.get("senha", "")
        ).strip()

        ok_u = hmac.compare_digest(usuario.encode(), SISTEMA_USUARIO.encode())
        ok_s = hmac.compare_digest(senha.encode(), SISTEMA_SENHA.encode())
        if ok_u and ok_s:
            tentativas_login.pop(ip, None)
            session["logado_sistema"] = True
            session.permanent = True
            registrar_auditoria(
                "Governança", "Login", f"Acesso à Central por IP {ip}"
            )
            flash("Login realizado com sucesso!")
            if eh_ajax:
                return jsonify(
                    {"ok": True, "redirect": url_for("central_gestao")}
                )
            return redirect(url_for("central_gestao"))

        tentativas_login[ip] = (falhas + 1, desde)
        if eh_ajax:
            return (
                jsonify({"ok": False, "erro": "Usuário ou senha incorretos."}),
                401,
            )
        flash("Usuário ou senha incorretos.")

    ultimo_os_id = db.session.query(db.func.max(OrdemServico.id)).scalar() or 0
    return renderizar(HTML_LOGIN_SISTEMA, ultimo_os_id=ultimo_os_id)


@app.route("/logout")
def logout_sistema():
    session.pop("logado_sistema", None)
    flash("Sessão encerrada com sucesso.")
    return redirect(url_for("lobby"))


@app.route("/health")
def healthcheck():
    return jsonify({"status": "ok", "timestamp": agora_formatado()}), 200


@app.route("/os/evidencia/<path:filename>")
def ver_evidencia_os(filename):
    nome_limpo = os.path.basename(filename)
    ext = os.path.splitext(nome_limpo)[1].lower()
    if ext not in EXTENSOES_EVIDENCIA:
        return "Arquivo inválido", 400
    caminho = os.path.join(EVIDENCIAS_DIR, nome_limpo)
    if os.path.isfile(caminho):
        return send_from_directory(
            EVIDENCIAS_DIR, nome_limpo, mimetype=EXTENSOES_EVIDENCIA[ext]
        )
    ev_db = EvidenciaOS.query.filter_by(nome_arquivo=nome_limpo).first()
    if ev_db and ev_db.dados_binarios:
        try:
            os.makedirs(EVIDENCIAS_DIR, exist_ok=True)
            with open(caminho, "wb") as f:
                f.write(ev_db.dados_binarios)
        except OSError:
            pass
        mime = ev_db.mime_type or EXTENSOES_EVIDENCIA[ext]
        return Response(ev_db.dados_binarios, mimetype=mime)
    return "Evidência não encontrada", 404


@app.route("/os/evidencia-atualizar/<int:id>", methods=["POST"])
def atualizar_evidencia_os(id):
    ordem = db.session.get(OrdemServico, id)
    if not ordem or ordem.status != "Finalizada":
        flash("Ordem de Serviço inválida ou ainda não finalizada.")
        return redirect(url_for("dashboard"))

    existentes = listar_evidencias_ordem(ordem)
    arquivos = request.files.getlist("evidencia_foto")
    novos, erro_ev = salvar_multiplas_evidencias_os(
        arquivos,
        ordem.numero_os,
        existentes=len(existentes),
        ordem_id=ordem.id,
    )
    if erro_ev:
        db.session.rollback()
        flash(erro_ev)
        return redirect(url_for("dashboard"))
    if not novos:
        flash("Selecione pelo menos uma imagem para anexar como evidência.")
        return redirect(url_for("dashboard"))

    try:
        todos = existentes + novos
        gravar_evidencias_ordem(ordem, todos)
        db.session.commit()
        registrar_auditoria(
            "Ordens de Serviço",
            "Evidências O.S. Adicionadas",
            f"O.S. {ordem.numero_os} recebeu +{len(novos)} foto(s) (total: {len(todos)})",
        )
        flash(
            f"{len(novos)} foto(s) adicionada(s) à O.S. {ordem.numero_os}!"
        )
    except Exception:
        db.session.rollback()
        for item in novos:
            remover_evidencia_os(item["arquivo"])
        flash("Erro ao salvar as evidências fotográficas.")
    return redirect(url_for("dashboard"))


@app.route("/os/evidencia-remover-item/<int:id>", methods=["POST"])
def remover_evidencia_item_rota(id):
    ordem = db.session.get(OrdemServico, id)
    if not ordem:
        flash("Ordem de Serviço não encontrada.")
        return redirect(url_for("dashboard"))

    alvo = os.path.basename(request.form.get("arquivo", "").strip())
    if not alvo:
        flash("Foto inválida.")
        return redirect(url_for("dashboard"))

    existentes = listar_evidencias_ordem(ordem)
    restantes = [item for item in existentes if item["arquivo"] != alvo]
    if len(restantes) == len(existentes):
        flash("Foto não encontrada nesta O.S.")
        return redirect(url_for("dashboard"))

    try:
        gravar_evidencias_ordem(ordem, restantes)
        db.session.commit()
        remover_evidencia_os(alvo)
        registrar_auditoria(
            "Ordens de Serviço",
            "Foto Evidência Removida",
            f"Foto removida da O.S. {ordem.numero_os} ({len(restantes)} restante(s))",
        )
        flash(f"Foto removida da O.S. {ordem.numero_os}.")
    except Exception:
        db.session.rollback()
        flash("Erro ao remover a foto.")
    return redirect(url_for("dashboard"))


@app.route("/os/evidencia-remover/<int:id>", methods=["POST"])
def remover_evidencia_rota(id):
    ordem = db.session.get(OrdemServico, id)
    if not ordem:
        flash("Ordem de Serviço não encontrada.")
        return redirect(url_for("dashboard"))

    existentes = listar_evidencias_ordem(ordem)
    if existentes:
        try:
            gravar_evidencias_ordem(ordem, [])
            db.session.commit()
            for item in existentes:
                remover_evidencia_os(item["arquivo"])
            registrar_auditoria(
                "Ordens de Serviço",
                "Evidências O.S. Removidas",
                f"Todas as {len(existentes)} foto(s) removidas da O.S. {ordem.numero_os}",
            )
            flash(f"Fotos de evidência removidas da O.S. {ordem.numero_os}.")
        except Exception:
            db.session.rollback()
            flash("Erro ao remover evidências.")
    return redirect(url_for("dashboard"))


@app.route("/api/os/novas")
def api_novas_os():
    try:
        ultimo_id = int(request.args.get("ultimo_id", 0))
    except (ValueError, TypeError):
        ultimo_id = 0

    max_id = db.session.query(db.func.max(OrdemServico.id)).scalar() or 0
    if ultimo_id <= 0:
        return jsonify({"ultimo_id": max_id, "ordens": []})

    novas = (
        OrdemServico.query.filter(OrdemServico.id > ultimo_id)
        .order_by(OrdemServico.id.asc())
        .limit(10)
        .all()
    )
    lista = [
        {
            "id": o.id,
            "numero_os": o.numero_os,
            "solicitante": o.solicitante,
            "setor": o.setor or "",
            "titulo": o.titulo or "",
            "prioridade": o.prioridade or "Média",
            "responsavel": o.responsavel or "",
            "abertura": o.abertura,
            "status": o.status,
        }
        for o in novas
    ]
    return jsonify({"ultimo_id": max_id, "ordens": lista})


# ==============================================================================
# ROTAS LOBBY E ORDENS DE SERVIÇO (CHÃO DE FÁBRICA)
# ==============================================================================
@app.route("/")
def lobby():
    ordens = OrdemServico.query.all()
    total = len(ordens)
    abertas = sum(1 for o in ordens if o.status == "Aberta")
    aguardando = sum(1 for o in ordens if o.status == "Aguardando Peça")
    ultimo_os_id = max((o.id for o in ordens), default=0)
    return renderizar(
        HTML_LOBBY,
        total=total,
        abertas=abertas,
        aguardando=aguardando,
        proxima=gerar_numero_os(),
        ultimo_os_id=ultimo_os_id,
    )


@app.route("/os")
def dashboard():
    busca = request.args.get("busca", "").strip()
    prefill_setor = request.args.get("abrir_maquina", "").strip()
    ordens = OrdemServico.query.order_by(OrdemServico.id.desc()).all()
    total = len(ordens)
    abertas = sum(1 for o in ordens if o.status == "Aberta")
    aguardando = sum(1 for o in ordens if o.status == "Aguardando Peça")
    finalizadas = sum(1 for o in ordens if o.status == "Finalizada")
    pct_final = round(finalizadas / total * 100) if total else 0
    ultimo_os_id = max((o.id for o in ordens), default=0)

    pecas_estoque = (
        ItemEstoque.query.filter(ItemEstoque.quantidade > 0)
        .order_by(ItemEstoque.nome.asc())
        .all()
    )
    maquinas_spare = [m.maquina for m in MaquinaSpare.query.all()]
    maquinas_os = [o.setor for o in ordens if o.setor]
    maquinas_sugeridas = sorted(list(set(maquinas_spare + maquinas_os)))

    return renderizar(
        HTML_OS,
        ordens=ordens,
        total=total,
        abertas=abertas,
        aguardando=aguardando,
        finalizadas=finalizadas,
        pct_final=pct_final,
        proxima=gerar_numero_os(),
        busca=busca,
        prefill_setor=prefill_setor,
        maquinas_sugeridas=maquinas_sugeridas,
        pecas_estoque=pecas_estoque,
        funcionarios_equipe=FUNCIONARIOS_EQUIPE,
        fmt_hhmm=horas_decimais_para_hhmm,
        fmt_input_time=horas_decimais_para_input_time,
        ultimo_os_id=ultimo_os_id,
    )


@app.route("/criar", methods=["POST"])
def criar():
    solicitante = request.form.get("solicitante", "").strip()
    if not solicitante:
        flash("Preencha o nome do solicitante.")
        return redirect(url_for("dashboard"))

    prioridade = request.form.get("prioridade", "Média")
    if prioridade not in PRIORIDADES_VALIDAS:
        prioridade = "Média"

    with os_lock:
        try:
            numero_os = gerar_numero_os()
            nova = OrdemServico(
                numero_os=numero_os,
                solicitante=solicitante[:100],
                setor=request.form.get("setor", "").strip()[:100],
                titulo=request.form.get("titulo", "").strip()[:200],
                descricao="",
                prioridade=prioridade,
                responsavel=request.form.get("responsavel", "").strip()[:100],
                abertura=agora_formatado(),
                status="Aberta",
            )
            db.session.add(nova)
            db.session.commit()
            registrar_auditoria(
                "Ordens de Serviço",
                "Abertura O.S.",
                f"O.S. {numero_os} aberta por {solicitante}",
            )
            flash(f"Ordem de Serviço {numero_os} registrada com sucesso.")
        except Exception:
            db.session.rollback()
            flash("Erro ao registrar a Ordem de Serviço. Tente novamente.")
    return redirect(url_for("dashboard"))


@app.route("/editar-os/<int:id>", methods=["POST"])
def editar_os(id):
    ordem = db.session.get(OrdemServico, id)
    if not ordem:
        flash("Ordem de Serviço não encontrada.")
        return redirect(url_for("dashboard"))

    solicitante = request.form.get("solicitante", "").strip()
    if not solicitante:
        flash("O campo Solicitante é obrigatório.")
        return redirect(url_for("dashboard"))

    prioridade = request.form.get("prioridade", "Média")
    if prioridade not in PRIORIDADES_VALIDAS:
        prioridade = "Média"

    status = request.form.get("status", "Aberta")
    if status not in STATUS_OS_VALIDOS:
        status = ordem.status

    tipo_manutencao = request.form.get("tipo_manutencao", "").strip()
    if tipo_manutencao not in TIPOS_MANUTENCAO:
        tipo_manutencao = ""

    funcs_selecionados = [
        f
        for f in request.form.getlist("funcionarios")
        if f in FUNCIONARIOS_VALIDOS
    ]
    tempo_str = request.form.get("tempo_trabalhado", "").strip()
    horas_decimais = (
        tempo_hhmm_para_horas_decimais(tempo_str)
        if tempo_str
        else (ordem.horas_trabalhadas or 0.0)
    )

    try:
        ordem.solicitante = solicitante[:100]
        ordem.setor = request.form.get("setor", "").strip()[:100]
        ordem.titulo = request.form.get("titulo", "").strip()[:200]
        ordem.responsavel = request.form.get("responsavel", "").strip()[:100]
        ordem.prioridade = prioridade
        ordem.status = status
        ordem.tipo_manutencao = tipo_manutencao
        ordem.funcionarios = ", ".join(funcs_selecionados)
        ordem.horas_trabalhadas = horas_decimais
        ordem.observacao_tecnica = request.form.get(
            "observacao_tecnica", ""
        ).strip()[:2000]

        if status == "Finalizada":
            if not ordem.finalizacao:
                ordem.finalizacao = agora_formatado()
            ordem.tempo_total = calcular_tempo_total(
                ordem.abertura, ordem.finalizacao
            )
            ordem.motivo_pausa = ""
        else:
            ordem.finalizacao = ""
            ordem.tempo_total = ""

        db.session.commit()
        registrar_auditoria(
            "Ordens de Serviço",
            "Edição O.S.",
            f"O.S. {ordem.numero_os} editada",
        )
        flash(f"Ordem de Serviço {ordem.numero_os} atualizada com sucesso.")
    except Exception:
        db.session.rollback()
        flash("Erro ao salvar alterações na Ordem de Serviço.")
    return redirect(url_for("dashboard"))


@app.route("/pausar/<int:id>", methods=["POST"])
def pausar_os(id):
    ordem = db.session.get(OrdemServico, id)
    if not ordem or ordem.status == "Finalizada":
        flash("Operação inválida para esta O.S.")
        return redirect(url_for("dashboard"))

    novo_status = request.form.get("novo_status", "Aguardando Peça")
    if novo_status not in ("Aguardando Peça", "Aberta"):
        novo_status = "Aguardando Peça"
    motivo = request.form.get("motivo_pausa", "").strip()[:200]

    ordem.status = novo_status
    ordem.motivo_pausa = motivo if novo_status == "Aguardando Peça" else ""
    db.session.commit()
    registrar_auditoria(
        "Ordens de Serviço",
        "Status O.S.",
        f"O.S. {ordem.numero_os} alterada para {novo_status}",
    )
    flash(f"Status da O.S. {ordem.numero_os} atualizado para {novo_status}.")
    return redirect(url_for("dashboard"))


@app.route("/finalizar/<int:id>", methods=["POST"])
def finalizar(id):
    ordem = db.session.get(OrdemServico, id)
    if not ordem or ordem.status == "Finalizada":
        flash("Ordem de Serviço inválida ou já finalizada.")
        return redirect(url_for("dashboard"))

    tipo_manutencao = request.form.get("tipo_manutencao", "").strip()
    if tipo_manutencao not in TIPOS_MANUTENCAO:
        flash("Selecione o Tipo de Manutenção (Preventiva ou Corretiva).")
        return redirect(url_for("dashboard"))

    funcs_selecionados = [
        f
        for f in request.form.getlist("funcionarios")
        if f in FUNCIONARIOS_VALIDOS
    ]
    if not funcs_selecionados:
        flash("Selecione pelo menos um funcionário que atuou nesta O.S.")
        return redirect(url_for("dashboard"))

    tempo_str = request.form.get("tempo_trabalhado", "").strip()
    horas_decimais = tempo_hhmm_para_horas_decimais(tempo_str)
    if horas_decimais <= 0:
        flash("Informe um tempo trabalhado válido (maior que 00:00).")
        return redirect(url_for("dashboard"))

    obs_tecnica = request.form.get("observacao_tecnica", "").strip()[:2000]
    peca_id_raw = request.form.get("peca_id", "").strip()
    peca_qtd_raw = request.form.get("peca_qtd", "1").strip()

    existentes = listar_evidencias_ordem(ordem)
    arquivos_evidencia = request.files.getlist("evidencia_foto")
    novos_itens_ev, erro_evidencia = salvar_multiplas_evidencias_os(
        arquivos_evidencia,
        ordem.numero_os,
        existentes=len(existentes),
        ordem_id=ordem.id,
    )
    if erro_evidencia:
        db.session.rollback()
        flash(erro_evidencia)
        return redirect(url_for("dashboard"))

    try:
        if peca_id_raw:
            try:
                p_id = int(peca_id_raw)
                p_qtd = int(peca_qtd_raw)
                peca = db.session.get(ItemEstoque, p_id)
                if peca and p_qtd > 0:
                    if peca.quantidade < p_qtd:
                        for item in novos_itens_ev:
                            remover_evidencia_os(item["arquivo"])
                        flash(
                            f"Saldo insuficiente em estoque para '{peca.nome}'"
                            f" (Saldo: {peca.quantidade})."
                        )
                        return redirect(url_for("dashboard"))
                    peca.quantidade -= p_qtd
                    peca.ultima_atualizacao = agora_formatado()
                    db.session.add(
                        HistoricoMovimentacao(
                            peca_codigo=peca.codigo,
                            peca_nome=peca.nome,
                            tipo="Retirada",
                            quantidade=p_qtd,
                            solicitante=f"Baixa via O.S. {ordem.numero_os}",
                            data_hora=agora_formatado(),
                        )
                    )
            except ValueError:
                pass

        ordem.finalizacao = agora_formatado()
        ordem.status = "Finalizada"
        ordem.motivo_pausa = ""
        ordem.tempo_total = calcular_tempo_total(
            ordem.abertura, ordem.finalizacao
        )
        ordem.tipo_manutencao = tipo_manutencao
        ordem.funcionarios = ", ".join(funcs_selecionados)
        ordem.horas_trabalhadas = horas_decimais
        ordem.observacao_tecnica = obs_tecnica
        if novos_itens_ev:
            gravar_evidencias_ordem(ordem, existentes + novos_itens_ev)

        db.session.commit()
        registrar_auditoria(
            "Ordens de Serviço",
            "Finalização O.S.",
            f"O.S. {ordem.numero_os} finalizada por {ordem.funcionarios} com"
            f" {len(novos_itens_ev)} foto(s)",
        )
        flash(
            f"Ordem de Serviço {ordem.numero_os} finalizada e auditada com"
            " sucesso!"
        )
    except Exception:
        db.session.rollback()
        for item in novos_itens_ev:
            remover_evidencia_os(item["arquivo"])
        flash("Erro ao finalizar Ordem de Serviço.")
    return redirect(url_for("dashboard"))


@app.route("/excluir/<int:id>", methods=["POST"])
def excluir(id):
    ordem = db.session.get(OrdemServico, id)
    if ordem:
        num = ordem.numero_os
        ev_antiga = ordem.evidencia_imagem
        db.session.delete(ordem)
        db.session.commit()
        remover_evidencia_os(ev_antiga)
        registrar_auditoria(
            "Ordens de Serviço", "Exclusão O.S.", f"O.S. {num} removida"
        )
        flash(f"Ordem de Serviço {num} excluída.")
    return redirect(url_for("dashboard"))


@app.route("/imprimir/os/<int:id>")
def imprimir_os(id):
    ordem = db.session.get(OrdemServico, id)
    if not ordem:
        return "Ordem de Serviço não encontrada.", 404
    return renderizar(
        HTML_IMPRESSAO_OS, os=ordem, fmt_hhmm=horas_decimais_para_hhmm
    )


@app.route("/excel")
def excel():
    ordens = OrdemServico.query.order_by(OrdemServico.id.desc()).all()
    dados = [
        {
            "Número O.S.": o.numero_os,
            "Solicitante": o.solicitante,
            "Máquina / Setor": o.setor or "",
            "Causa": o.titulo or "",
            "Prioridade": o.prioridade,
            "Responsável": o.responsavel or "",
            "Tipo Manutenção": o.tipo_manutencao or "",
            "Funcionários": o.funcionarios or "",
            "Tempo Apontado": (
                horas_decimais_para_hhmm(o.horas_trabalhadas)
                if o.horas_trabalhadas
                else ""
            ),
            "Laudo Técnico": o.observacao_tecnica or "",
            "Evidências Fotográficas": (
                f"{len(o.evidencias_lista)} foto(s)"
                if o.evidencias_lista
                else ""
            ),
            "Status": o.status,
            "Abertura": o.abertura,
            "Finalização": o.finalizacao or "",
            "Lead Time": o.tempo_total or "",
        }
        for o in ordens
    ]
    return gerar_excel(
        pd.DataFrame(dados),
        "Ordens de Servico",
        "ordens_servico",
        coluna_status="Status",
    )


# ==============================================================================
# CENTRAL DE GESTÃO UNIFICADA (PROTEGIDA POR SENHA)
# ==============================================================================
@app.route("/central-gestao")
@requer_login
def central_gestao():
    pecas = ItemEstoque.query.order_by(ItemEstoque.nome.asc()).all()
    movimentacoes = (
        HistoricoMovimentacao.query.order_by(HistoricoMovimentacao.id.desc())
        .limit(40)
        .all()
    )
    total_pecas = len(pecas)
    total_quantidade = sum(p.quantidade for p in pecas)
    itens_baixo = sum(1 for p in pecas if p.quantidade <= p.estoque_minimo)
    total_movimentacoes = HistoricoMovimentacao.query.count()
    mapa_estoque = {p.codigo.upper(): p.quantidade for p in pecas}

    maquinas = MaquinaSpare.query.order_by(MaquinaSpare.maquina.asc()).all()

    hoje = date.today()
    planos = PlanoPreventivo.query.all()
    for p in planos:
        p.status_calc = status_plano(p, hoje)
    planos.sort(
        key=lambda x: (
            0 if x.status_calc == "Atrasado" else 1,
            parse_data_br(x.proxima_execucao) or date.max,
        )
    )
    total_planos = len(planos)
    atrasados = sum(1 for p in planos if p.status_calc == "Atrasado")
    agendados = total_planos - atrasados
    arquivos_pcm = ArquivoPCM.query.order_by(ArquivoPCM.id.desc()).all()

    compras = SolicitacaoCompra.query.order_by(
        SolicitacaoCompra.id.desc()
    ).all()
    total_sc = len(compras)
    emergenciais = sum(
        1 for c in compras if c.tipo_sc == "S.C. Emergencial"
    )
    programadas = sum(
        1 for c in compras if c.tipo_sc == "S.C. Programada"
    )
    reposicao = sum(
        1 for c in compras if c.tipo_sc == "S.C. Reposição / Estoque"
    )

    repor_codigo = request.args.get("repor_codigo", "").strip()
    repor_nome = request.args.get("repor_nome", "").strip()
    repor_qtd = request.args.get("repor_qtd", "").strip()
    repor_item = (
        {"codigo": repor_codigo, "nome": repor_nome, "qtd": repor_qtd}
        if repor_codigo
        else None
    )

    ordens_fin = OrdemServico.query.filter_by(status="Finalizada").all()
    total_fin = len(ordens_fin)
    total_preventiva = sum(
        1 for o in ordens_fin if o.tipo_manutencao == "Preventiva"
    )
    total_corretiva = sum(
        1 for o in ordens_fin if o.tipo_manutencao == "Corretiva"
    )
    pct_geral_prev = (
        round((total_preventiva / total_fin) * 100, 1) if total_fin > 0 else 0
    )

    corretivas = [
        o for o in ordens_fin if o.tipo_manutencao == "Corretiva"
    ]
    horas_corr = [
        o.horas_trabalhadas
        for o in corretivas
        if (o.horas_trabalhadas or 0) > 0
    ]
    mttr_horas = sum(horas_corr) / len(horas_corr) if horas_corr else 0.0
    mttr_str = horas_decimais_para_hhmm(mttr_horas)

    grupos_maquinas = {}
    for o in corretivas:
        nome_maq = (o.setor or "").strip() or "Sem Máquina Informada"
        chave_maq = normalizar_chave_maquina(nome_maq)
        if chave_maq not in grupos_maquinas:
            grupos_maquinas[chave_maq] = {
                "chave": chave_maq,
                "nome": nome_maq,
                "datas": [],
                "horas": [],
                "corretivas": 0,
            }
        grupos_maquinas[chave_maq]["corretivas"] += 1
        if (o.horas_trabalhadas or 0) > 0:
            grupos_maquinas[chave_maq]["horas"].append(o.horas_trabalhadas)
        try:
            dt_ab = datetime.strptime(o.abertura, FORMATO_DATA)
            grupos_maquinas[chave_maq]["datas"].append(dt_ab)
        except Exception:
            pass

    mtbf_maquinas = []
    intervalos_gerais_horas = []
    for chave_maq, info in grupos_maquinas.items():
        datas_ord = sorted(info["datas"])
        mtbf_maq_h = 0.0
        if len(datas_ord) >= 2:
            diffs_h = [
                max(0.0, (datas_ord[i] - datas_ord[i - 1]).total_seconds() / 3600.0)
                for i in range(1, len(datas_ord))
            ]
            mtbf_maq_h = sum(diffs_h) / len(diffs_h)
            intervalos_gerais_horas.extend(diffs_h)

        mttr_maq_h = (
            sum(info["horas"]) / len(info["horas"]) if info["horas"] else 0.0
        )
        horas_totais_maq = sum(info["horas"])
        mtbf_maquinas.append({
            "chave": chave_maq,
            "nome": info["nome"],
            "corretivas": info["corretivas"],
            "mtbf_horas": round(mtbf_maq_h, 2),
            "mtbf_str": (
                formatar_duracao_indicador(mtbf_maq_h)
                if len(datas_ord) >= 2
                else "Aguard. 2ª falha"
            ),
            "mttr_str": horas_decimais_para_hhmm(mttr_maq_h),
            "horas_str": horas_decimais_para_hhmm(horas_totais_maq),
        })

    mtbf_maquinas.sort(key=lambda x: (-x["corretivas"], x["nome"].lower()))
    mtbf_horas = (
        sum(intervalos_gerais_horas) / len(intervalos_gerais_horas)
        if intervalos_gerais_horas
        else 0.0
    )
    mtbf_str = (
        formatar_duracao_indicador(mtbf_horas)
        if intervalos_gerais_horas
        else "Sem recorrência"
    )

    agrupamento = {}
    for o in ordens_fin:
        chave = chave_mes_os(o) or "Sem Data"
        if chave not in agrupamento:
            agrupamento[chave] = {
                "prev": 0,
                "corr": 0,
                "total": 0,
                "horas": 0.0,
            }
        agrupamento[chave]["total"] += 1
        agrupamento[chave]["horas"] += o.horas_trabalhadas or 0.0
        if o.tipo_manutencao == "Preventiva":
            agrupamento[chave]["prev"] += 1
        elif o.tipo_manutencao == "Corretiva":
            agrupamento[chave]["corr"] += 1

    dados_tabela = []
    meses_lista = []
    for chave in sorted(agrupamento.keys(), reverse=True):
        item = agrupamento[chave]
        partes = chave.split("-")
        mes_nome = (
            f"{MESES_PT.get(partes[1], partes[1])}/{partes[0]}"
            if len(partes) == 2
            else chave
        )
        meses_lista.append({"chave": chave, "nome": mes_nome})
        dados_tabela.append({
            "chave": chave,
            "mes_nome": mes_nome,
            "total_os": item["total"],
            "prev_qtd": item["prev"],
            "corr_qtd": item["corr"],
        })

    mes_sel = request.args.get("mes", "geral")
    if mes_sel != "geral":
        ordens_filtradas = [
            o for o in ordens_fin if chave_mes_os(o) == mes_sel
        ]
    else:
        ordens_filtradas = ordens_fin

    funcionarios_indiv = []
    for f_info in FUNCIONARIOS_EQUIPE:
        valor_id = f_info["valor"]
        nome_curto = f_info["nome"]
        horas_totais = 0.0
        qtd_os = 0
        dias_unicos = set()

        for o in ordens_filtradas:
            lista_funcs = [
                x.strip()
                for x in (o.funcionarios or "").split(",")
                if x.strip()
            ]
            # Compatibilidade com registros antigos salvos apenas como "Jefferson"
            match_legado = (
                nome_curto == "Jefferson"
                and "Jefferson" in lista_funcs
                and f_info["funcao"] == "Eletricista"
            )
            if valor_id in lista_funcs or match_legado:
                qtd_os += 1
                horas_totais += o.horas_trabalhadas or 0.0
                for campo_dt in (o.finalizacao, o.abertura):
                    try:
                        dt_obj = datetime.strptime(campo_dt, FORMATO_DATA)
                        dias_unicos.add(dt_obj.date())
                        break
                    except Exception:
                        continue

        qtd_dias = len(dias_unicos)
        capacidade = qtd_dias * JORNADA_REFERENCIA
        ef_indiv = (
            round((horas_totais / capacidade) * 100, 1)
            if capacidade > 0
            else 0.0
        )

        funcionarios_indiv.append({
            "nome": f_info["nome"],
            "funcao": f_info["funcao"],
            "icone": f_info["icone"],
            "equipe": f_info["equipe"],
            "turno": f_info["turno"],
            "horas_totais": round(horas_totais, 2),
            "tempo_str": horas_decimais_para_hhmm(horas_totais),
            "qtd_os": qtd_os,
            "dias": qtd_dias,
            "eficiencia": ef_indiv,
            "eficiencia_clamp": min(100, ef_indiv),
        })

    logs = (
        LogAuditoria.query.order_by(LogAuditoria.id.desc()).limit(60).all()
    )
    tamanho_db = (
        round(os.path.getsize(DB_PATH) / 1024, 1)
        if not IS_POSTGRES and os.path.exists(DB_PATH)
        else 0
    )
    ultimo_os_id = db.session.query(db.func.max(OrdemServico.id)).scalar() or 0
    resumo_banco = {
        "ordens": OrdemServico.query.count(),
        "estoque": total_pecas,
        "maquinas": len(maquinas),
        "planos": total_planos,
        "compras": total_sc,
    }

    return renderizar(
        HTML_CENTRAL_GESTAO,
        pecas=pecas,
        movimentacoes=movimentacoes,
        total_pecas=total_pecas,
        total_quantidade=total_quantidade,
        itens_baixo=itens_baixo,
        total_movimentacoes=total_movimentacoes,
        mapa_estoque=mapa_estoque,
        maquinas=maquinas,
        planos=planos,
        total_planos=total_planos,
        agendados=agendados,
        atrasados=atrasados,
        arquivos_pcm=arquivos_pcm,
        compras=compras,
        total_sc=total_sc,
        emergenciais=emergenciais,
        programadas=programadas,
        reposicao=reposicao,
        proxima_sc=gerar_numero_sc(),
        hoje=hoje_iso(),
        centros_custo_opcoes=CENTROS_CUSTO_OFICIAIS,
        repor_item=repor_item,
        total_fin=total_fin,
        total_preventiva=total_preventiva,
        total_corretiva=total_corretiva,
        pct_geral_prev=pct_geral_prev,
        mttr_str=mttr_str,
        mtbf_str=mtbf_str,
        mtbf_maquinas=mtbf_maquinas,
        dados_tabela=dados_tabela,
        total_meses=len(dados_tabela),
        meses=meses_lista,
        mes_sel=mes_sel,
        funcionarios_indiv=funcionarios_indiv,
        logs=logs,
        tamanho_db=tamanho_db,
        is_postgres=IS_POSTGRES,
        resumo_banco=resumo_banco,
        ultimo_os_id=ultimo_os_id,
    )


# ==============================================================================
# ROTAS DE ESTOQUE, SPARE PARTS, PCM, COMPRAS & GOVERNANÇA
# ==============================================================================
@app.route("/estoque/cadastrar", methods=["POST"])
@requer_login
def cadastrar_peca():
    codigo = request.form.get("codigo", "").strip().upper()
    nome = request.form.get("nome", "").strip()
    if not codigo or not nome:
        flash("Preencha o Código e o Nome da Peça.")
        return redirect(url_for("central_gestao", tab="estoque"))
    try:
        quantidade = max(0, int(request.form.get("quantidade", 0)))
        estoque_minimo = max(1, int(request.form.get("estoque_minimo", 2)))
    except ValueError:
        flash("Valores inválidos.")
        return redirect(url_for("central_gestao", tab="estoque"))

    with estoque_lock:
        if ItemEstoque.query.filter_by(codigo=codigo).first():
            flash(f"Já existe uma peça com o código {codigo}.")
            return redirect(url_for("central_gestao", tab="estoque"))
        peca = ItemEstoque(
            codigo=codigo[:50],
            nome=nome[:150],
            categoria=request.form.get("categoria", "").strip()[:100]
            or "Geral",
            localizacao=request.form.get("localizacao", "").strip()[:100],
            quantidade=quantidade,
            estoque_minimo=estoque_minimo,
            ultima_atualizacao=agora_formatado(),
        )
        db.session.add(peca)
        if quantidade > 0:
            db.session.add(
                HistoricoMovimentacao(
                    peca_codigo=codigo[:50],
                    peca_nome=nome[:150],
                    tipo="Entrada",
                    quantidade=quantidade,
                    solicitante="Cadastro Inicial",
                    data_hora=agora_formatado(),
                )
            )
        db.session.commit()
        registrar_auditoria(
            "Estoque",
            "Cadastro Peça",
            f"SKU {codigo} - {nome} (Qtd: {quantidade})",
        )
        flash(f"Peça '{nome}' cadastrada com sucesso.")
    return redirect(url_for("central_gestao", tab="estoque"))


@app.route("/estoque/movimentar/<int:id>", methods=["POST"])
@requer_login
def movimentar_estoque(id):
    tipo = request.form.get("tipo", "Retirada")
    solicitante = request.form.get("solicitante", "").strip()
    try:
        qtd = int(request.form.get("qtd", 0))
    except ValueError:
        qtd = 0
    if qtd <= 0:
        flash("Informe quantidade maior que zero.")
        return redirect(url_for("central_gestao", tab="estoque"))

    with estoque_lock:
        peca = db.session.get(ItemEstoque, id)
        if not peca:
            return redirect(url_for("central_gestao", tab="estoque"))
        if tipo == "Retirada":
            if peca.quantidade < qtd:
                flash(f"Saldo insuficiente ({peca.quantidade}).")
                return redirect(url_for("central_gestao", tab="estoque"))
            peca.quantidade -= qtd
            origem = solicitante[:100]
        else:
            peca.quantidade += qtd
            origem = solicitante[:100] or "Reposição de Estoque"
        peca.ultima_atualizacao = agora_formatado()
        db.session.add(
            HistoricoMovimentacao(
                peca_codigo=peca.codigo,
                peca_nome=peca.nome,
                tipo=tipo,
                quantidade=qtd,
                solicitante=origem,
                data_hora=agora_formatado(),
            )
        )
        db.session.commit()
        registrar_auditoria(
            "Estoque", tipo, f"{qtd}x {peca.codigo} ({origem})"
        )
        flash(f"{tipo} de {qtd} unidade(s) registrada.")
    return redirect(url_for("central_gestao", tab="estoque"))


@app.route("/estoque/editar/<int:id>", methods=["POST"])
@requer_login
def editar_peca(id):
    peca = db.session.get(ItemEstoque, id)
    if not peca:
        return redirect(url_for("central_gestao", tab="estoque"))
    codigo = request.form.get("codigo", "").strip().upper()
    nome = request.form.get("nome", "").strip()
    try:
        estoque_minimo = max(1, int(request.form.get("estoque_minimo", 2)))
    except ValueError:
        return redirect(url_for("central_gestao", tab="estoque"))

    peca.codigo = codigo[:50]
    peca.nome = nome[:150]
    peca.categoria = (
        request.form.get("categoria", "").strip()[:100] or "Geral"
    )
    peca.localizacao = request.form.get("localizacao", "").strip()[:100]
    peca.estoque_minimo = estoque_minimo
    peca.ultima_atualizacao = agora_formatado()
    db.session.commit()
    registrar_auditoria("Estoque", "Edição Peça", f"SKU {codigo} atualizado")
    flash(f"Peça '{peca.nome}' atualizada.")
    return redirect(url_for("central_gestao", tab="estoque"))


@app.route("/estoque/excluir/<int:id>", methods=["POST"])
@requer_login
def excluir_peca(id):
    peca = db.session.get(ItemEstoque, id)
    if peca:
        cod = peca.codigo
        db.session.delete(peca)
        db.session.commit()
        registrar_auditoria("Estoque", "Exclusão Peça", f"SKU {cod} removido")
        flash(f"Peça '{peca.nome}' excluída.")
    return redirect(url_for("central_gestao", tab="estoque"))


@app.route("/estoque/excel")
@requer_login
def excel_estoque():
    pecas = ItemEstoque.query.order_by(ItemEstoque.nome.asc()).all()
    dados = [
        {
            "Código": p.codigo,
            "Peça": p.nome,
            "Categoria": p.categoria or "",
            "Localização": p.localizacao or "",
            "Saldo Atual": p.quantidade,
            "Estoque Mínimo": p.estoque_minimo,
            "Status": (
                "Baixo Estoque"
                if p.quantidade <= p.estoque_minimo
                else "Normal"
            ),
            "Última Atualização": p.ultima_atualizacao,
        }
        for p in pecas
    ]
    return gerar_excel(
        pd.DataFrame(dados),
        "Estoque",
        "estoque_manutencao",
        coluna_status="Status",
    )


@app.route("/spare-parts/cadastrar-maquina", methods=["POST"])
@requer_login
def cadastrar_maquina_spare():
    maquina = request.form.get("maquina", "").strip()
    modelo = request.form.get("modelo", "").strip()
    if maquina and modelo:
        db.session.add(
            MaquinaSpare(
                maquina=maquina[:150],
                modelo=modelo[:100],
                setor=request.form.get("setor", "").strip()[:100],
                pecas_necessarias=request.form.get(
                    "pecas_necessarias", ""
                ).strip(),
            )
        )
        db.session.commit()
        registrar_auditoria(
            "Spare Parts", "Cadastro Máquina", f"{maquina} ({modelo})"
        )
        flash(f"Equipamento '{maquina}' cadastrado!")
    return redirect(url_for("central_gestao", tab="spare"))


@app.route("/spare-parts/editar-maquina/<int:id>", methods=["POST"])
@requer_login
def editar_maquina_spare(id):
    maq = db.session.get(MaquinaSpare, id)
    if not maq:
        flash("Equipamento não encontrado.")
        return redirect(url_for("central_gestao", tab="spare"))

    maquina = request.form.get("maquina", "").strip()
    modelo = request.form.get("modelo", "").strip()
    if not maquina or not modelo:
        flash("Preencha o nome da máquina e o modelo.")
        return redirect(url_for("central_gestao", tab="spare"))

    maq.maquina = maquina[:150]
    maq.modelo = modelo[:100]
    maq.setor = request.form.get("setor", "").strip()[:100]
    maq.pecas_necessarias = request.form.get("pecas_necessarias", "").strip()
    db.session.commit()
    registrar_auditoria(
        "Spare Parts", "Edição Máquina", f"{maq.maquina} ({maq.modelo})"
    )
    flash(f"Equipamento '{maq.maquina}' atualizado com sucesso!")
    return redirect(url_for("central_gestao", tab="spare"))


@app.route("/spare-parts/adicionar-peca/<int:id>", methods=["POST"])
@requer_login
def adicionar_peca_spare(id):
    maq = db.session.get(MaquinaSpare, id)
    nome_peca = request.form.get("nome_peca", "").strip()
    if maq and nome_peca:
        try:
            qtd = max(1, int(request.form.get("quantidade", 1)))
        except ValueError:
            qtd = 1
        db.session.add(
            PecaSpare(
                maquina_spare_id=maq.id,
                nome_peca=nome_peca[:150],
                codigo_peca=request.form.get("codigo_peca", "")
                .strip()
                .upper()[:50],
                quantidade=qtd,
                observacao=request.form.get("observacao", "").strip()[:200],
            )
        )
        db.session.commit()
        flash(f"Peça '{nome_peca}' vinculada!")
    return redirect(url_for("central_gestao", tab="spare"))


@app.route("/spare-parts/excluir-peca/<int:id>", methods=["POST"])
@requer_login
def excluir_peca_spare(id):
    peca = db.session.get(PecaSpare, id)
    if peca:
        db.session.delete(peca)
        db.session.commit()
        flash("Peça removida.")
    return redirect(url_for("central_gestao", tab="spare"))


@app.route("/spare-parts/excluir-maquina/<int:id>", methods=["POST"])
@requer_login
def excluir_maquina_spare(id):
    maq = db.session.get(MaquinaSpare, id)
    if maq:
        db.session.delete(maq)
        db.session.commit()
        flash("Máquina excluída.")
    return redirect(url_for("central_gestao", tab="spare"))


@app.route("/spare-parts/etiquetas")
@requer_login
def etiquetas_qrcode():
    maquinas = MaquinaSpare.query.order_by(MaquinaSpare.maquina.asc()).all()
    return renderizar(HTML_QRCODES, maquinas=maquinas)


@app.route("/pcm/cadastrar", methods=["POST"])
@requer_login
def cadastrar_plano_pcm():
    maquina = request.form.get("maquina", "").strip()
    atividade = request.form.get("atividade", "").strip()
    proxima = parse_data_br(request.form.get("proxima_execucao"))
    if maquina and atividade and proxima:
        freq = request.form.get("frequencia", "Mensal")
        if freq not in FREQUENCIA_PCM:
            freq = "Mensal"
        db.session.add(
            PlanoPreventivo(
                maquina=maquina[:150],
                codigo_maquina=request.form.get("codigo_maquina", "")
                .strip()
                .upper()[:50],
                setor=request.form.get("setor", "").strip()[:100],
                atividade=atividade[:2000],
                frequencia=freq,
                responsavel=request.form.get("responsavel", "").strip()[:100],
                proxima_execucao=proxima.strftime("%d/%m/%Y"),
                status="Agendado",
            )
        )
        db.session.commit()
        registrar_auditoria(
            "PCM", "Novo Plano", f"Preventivo cadastrado para {maquina}"
        )
        flash("Plano preventivo registrado!")
    return redirect(url_for("central_gestao", tab="pcm"))


@app.route("/pcm/importar-excel", methods=["POST"])
@requer_login
def importar_excel_pcm():
    arquivo = request.files.get("arquivo_excel")
    if not arquivo or not (arquivo.filename or "").strip():
        flash("Selecione um arquivo Excel (.xlsx ou .xls).")
        return redirect(url_for("central_gestao", tab="pcm"))

    nome_orig = os.path.basename(arquivo.filename).strip()
    ext = os.path.splitext(nome_orig)[1].lower()
    if ext not in EXTENSOES_PCM_EXCEL:
        flash("Formato inválido. Envie apenas planilhas Excel (.xlsx ou .xls).")
        return redirect(url_for("central_gestao", tab="pcm"))

    conteudo = arquivo.read()
    if not conteudo:
        flash("O arquivo selecionado está vazio.")
        return redirect(url_for("central_gestao", tab="pcm"))
    if len(conteudo) > TAMANHO_MAX_PCM_EXCEL:
        flash("O arquivo excede o limite de 25 MB.")
        return redirect(url_for("central_gestao", tab="pcm"))

    os.makedirs(PCM_FILES_DIR, exist_ok=True)
    slug = re.sub(
        r"[^A-Za-z0-9_-]", "_", os.path.splitext(nome_orig)[0]
    )[:50]
    nome_seguro = f"pcm_{slug}_{secrets.token_hex(5)}{ext}"
    destino = os.path.join(PCM_FILES_DIR, nome_seguro)

    try:
        with open(destino, "wb") as f:
            f.write(conteudo)

        reg = ArquivoPCM(
            nome_original=nome_orig[:200],
            nome_arquivo=nome_seguro,
            tamanho_kb=round(len(conteudo) / 1024, 1),
            enviado_por="PCM",
            data_envio=agora_formatado(),
        )
        db.session.add(reg)
        db.session.commit()
        registrar_auditoria(
            "PCM",
            "Arquivo Excel Salvo",
            f"Planilha '{nome_orig}' arquivada no PCM",
        )
        flash(
            f"Planilha '{nome_orig}' salva com sucesso para abertura no Excel!"
        )
    except Exception:
        db.session.rollback()
        if os.path.exists(destino):
            try:
                os.remove(destino)
            except OSError:
                pass
        flash("Erro ao guardar a planilha Excel no PCM.")
    return redirect(url_for("central_gestao", tab="pcm"))


@app.route("/pcm/arquivo/<int:id>/abrir")
@requer_login
def abrir_arquivo_pcm(id):
    reg = db.session.get(ArquivoPCM, id)
    if not reg:
        flash("Arquivo Excel não encontrado.")
        return redirect(url_for("central_gestao", tab="pcm"))
    nome_limpo = os.path.basename(reg.nome_arquivo)
    caminho = os.path.join(PCM_FILES_DIR, nome_limpo)
    if not os.path.isfile(caminho):
        flash("O arquivo físico não foi encontrado no servidor.")
        return redirect(url_for("central_gestao", tab="pcm"))
    ext = os.path.splitext(nome_limpo)[1].lower()
    mime = EXTENSOES_PCM_EXCEL.get(
        ext,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    return send_file(
        caminho,
        as_attachment=True,
        download_name=reg.nome_original,
        mimetype=mime,
    )


@app.route("/pcm/arquivo/<int:id>/excluir", methods=["POST"])
@requer_login
def excluir_arquivo_pcm(id):
    reg = db.session.get(ArquivoPCM, id)
    if reg:
        nome_orig = reg.nome_original
        caminho = os.path.join(
            PCM_FILES_DIR, os.path.basename(reg.nome_arquivo)
        )
        db.session.delete(reg)
        db.session.commit()
        if os.path.exists(caminho):
            try:
                os.remove(caminho)
            except OSError:
                pass
        registrar_auditoria(
            "PCM", "Arquivo Excel Excluído", f"Planilha '{nome_orig}' removida"
        )
        flash(f"Arquivo '{nome_orig}' removido.")
    return redirect(url_for("central_gestao", tab="pcm"))


@app.route("/pcm/gerar-os/<int:id>", methods=["POST"])
@requer_login
def gerar_os_pcm(id):
    plano = db.session.get(PlanoPreventivo, id)
    if not plano:
        return redirect(url_for("central_gestao", tab="pcm"))
    with os_lock:
        num = gerar_numero_os()
        nova_os = OrdemServico(
            numero_os=num,
            solicitante="PCM Automático",
            setor=f"{plano.maquina} ({plano.setor})"
            if plano.setor
            else plano.maquina,
            titulo=f"Preventiva {plano.frequencia} - {plano.maquina}",
            descricao=plano.atividade,
            prioridade="Média",
            responsavel=plano.responsavel or "Equipe PCM",
            abertura=agora_formatado(),
            status="Aberta",
            tipo_manutencao="Preventiva",
        )
        db.session.add(nova_os)
        db.session.commit()
        registrar_auditoria(
            "PCM",
            "Geração O.S. Preventiva",
            f"O.S. {num} gerada para {plano.maquina}",
        )
        flash(f"O.S. {num} gerada diretamente na fila de manutenção!")
    return redirect(url_for("central_gestao", tab="pcm"))


@app.route("/pcm/concluir/<int:id>", methods=["POST"])
@requer_login
def concluir_plano_pcm(id):
    plano = db.session.get(PlanoPreventivo, id)
    if plano:
        base = parse_data_br(plano.proxima_execucao) or date.today()
        plano.proxima_execucao = avancar_data(
            max(base, date.today()), plano.frequencia
        ).strftime("%d/%m/%Y")
        plano.status = "Agendado"
        db.session.commit()
        registrar_auditoria(
            "PCM",
            "Ciclo Concluído",
            f"{plano.maquina} reagendada para {plano.proxima_execucao}",
        )
        flash(
            f"Preventivo concluído! Próxima data: {plano.proxima_execucao}."
        )
    return redirect(url_for("central_gestao", tab="pcm"))


@app.route("/pcm/editar/<int:id>", methods=["POST"])
@requer_login
def editar_plano_pcm(id):
    plano = db.session.get(PlanoPreventivo, id)
    if plano:
        proxima = parse_data_br(request.form.get("proxima_execucao"))
        if proxima:
            plano.maquina = request.form.get("maquina", "").strip()[:150]
            plano.codigo_maquina = (
                request.form.get("codigo_maquina", "").strip().upper()[:50]
            )
            plano.setor = request.form.get("setor", "").strip()[:100]
            plano.frequencia = request.form.get("frequencia", "Mensal")
            plano.responsavel = request.form.get("responsavel", "").strip()[
                :100
            ]
            plano.atividade = request.form.get("atividade", "").strip()[:2000]
            plano.proxima_execucao = proxima.strftime("%d/%m/%Y")
            db.session.commit()
            flash("Planejamento atualizado!")
    return redirect(url_for("central_gestao", tab="pcm"))


@app.route("/pcm/excluir/<int:id>", methods=["POST"])
@requer_login
def excluir_plano_pcm(id):
    plano = db.session.get(PlanoPreventivo, id)
    if plano:
        db.session.delete(plano)
        db.session.commit()
        flash("Plano excluído.")
    return redirect(url_for("central_gestao", tab="pcm"))


@app.route("/pcm/excel")
@requer_login
def excel_pcm():
    planos = PlanoPreventivo.query.all()
    hoje = date.today()
    dados = [
        {
            "Máquina": p.maquina,
            "Tag": p.codigo_maquina or "",
            "Setor": p.setor or "",
            "Frequência": p.frequencia,
            "Próxima Execução": p.proxima_execucao,
            "Responsável": p.responsavel or "",
            "Status": status_plano(p, hoje),
            "Roteiro": p.atividade,
        }
        for p in planos
    ]
    return gerar_excel(
        pd.DataFrame(dados), "PCM", "cronograma_pcm", coluna_status="Status"
    )


@app.route("/compras/cadastrar", methods=["POST"])
@requer_login
def cadastrar_compras():
    solicitante = request.form.get("solicitante", "").strip()
    objeto_solicitacao = request.form.get("objeto_solicitacao", "").strip()
    descritivos = request.form.getlist("descritivo_material[]")
    quantidades = request.form.getlist("quantidade[]")
    unidades = request.form.getlist("unidade[]")
    links = request.form.getlist("link_referencia[]")

    if not solicitante or not objeto_solicitacao or not descritivos:
        flash("Preencha todos os campos obrigatórios da S.C.")
        return redirect(url_for("central_gestao", tab="compras"))

    tipo_sc = request.form.get("tipo_sc", "S.C. Emergencial").strip()
    if tipo_sc not in TIPOS_SC_VALIDOS:
        tipo_sc = "S.C. Emergencial"
    dt = parse_data_br(request.form.get("data_solicitacao")) or date.today()

    with sc_lock:
        num_sc = (
            request.form.get("numero_sc", "").strip() or gerar_numero_sc()
        )
        nova_sc = SolicitacaoCompra(
            numero_sc=num_sc[:100],
            tipo_sc=tipo_sc,
            solicitante=solicitante[:100],
            setor_solicitante=request.form.get(
                "setor_solicitante", "Manutenção"
            ).strip()[:100],
            responsaveis_duvidas=request.form.get(
                "responsaveis_duvidas", ""
            ).strip()[:150],
            centro_custo=request.form.get("centro_custo", "").strip()[:150],
            objeto_solicitacao=objeto_solicitacao[:250],
            data_solicitacao=dt.strftime("%d/%m/%Y"),
            status_aprovacao="Aguardando Aprovação",
        )
        for i in range(len(descritivos)):
            desc = descritivos[i].strip()
            if not desc:
                continue
            try:
                q = max(1, int(quantidades[i]))
            except (ValueError, IndexError):
                q = 1
            unid = unidades[i].strip() if i < len(unidades) else "Un"
            lnk = links[i].strip() if i < len(links) else ""
            nova_sc.itens.append(
                ItemSolicitacaoCompra(
                    descritivo=desc,
                    quantidade=q,
                    unidade=unid,
                    link_referencia=lnk,
                )
            )
        db.session.add(nova_sc)
        db.session.commit()
        registrar_auditoria(
            "Compras",
            "Nova S.C.",
            f"{num_sc} registrada por {solicitante} ({tipo_sc})",
        )
        flash(f"Solicitação {num_sc} registrada com sucesso!")
    return redirect(url_for("central_gestao", tab="compras"))


@app.route("/compras/atualizar-status/<int:id>", methods=["POST"])
@requer_login
def atualizar_status_sc(id):
    sc = db.session.get(SolicitacaoCompra, id)
    if not sc:
        return jsonify({"ok": False, "erro": "Não encontrada"}), 404
    novo_status = request.form.get("status_aprovacao", "").strip()
    if novo_status not in STATUS_SC_VALIDOS:
        return jsonify({"ok": False, "erro": "Status inválido"}), 400
    sc.status_aprovacao = novo_status
    db.session.commit()
    registrar_auditoria(
        "Compras",
        "Status S.C.",
        f"{sc.numero_sc} alterada para {novo_status}",
    )
    return jsonify({"ok": True})


@app.route("/compras/imprimir/<int:id>")
@requer_login
def imprimir_sc(id):
    sc = db.session.get(SolicitacaoCompra, id)
    if not sc:
        return "S.C. não encontrada", 404
    return renderizar(HTML_IMPRESSAO_SC, sc=sc)


@app.route("/compras/excluir/<int:id>", methods=["POST"])
@requer_login
def excluir_compras(id):
    sc = db.session.get(SolicitacaoCompra, id)
    if sc:
        num = sc.numero_sc
        db.session.delete(sc)
        db.session.commit()
        registrar_auditoria(
            "Compras", "Exclusão S.C.", f"{num} removida"
        )
        flash("Solicitação excluída.")
    return redirect(url_for("central_gestao", tab="compras"))


@app.route("/compras/excel")
@requer_login
def excel_compras():
    compras = SolicitacaoCompra.query.order_by(
        SolicitacaoCompra.id.desc()
    ).all()
    dados = []
    for c in compras:
        for idx, it in enumerate(c.itens, start=1):
            dados.append({
                "Número S.C.": c.numero_sc,
                "Classificação": c.tipo_sc,
                "Status": c.status_aprovacao,
                "Data": c.data_solicitacao,
                "Solicitante": c.solicitante,
                "Centro de Custo": c.centro_custo,
                "Objeto": c.objeto_solicitacao,
                "Item": idx,
                "Descritivo": it.descritivo,
                "Qtd": it.quantidade,
                "Unidade": it.unidade,
                "Link / Referência": it.link_referencia or "",
            })
    return gerar_excel(
        pd.DataFrame(dados),
        "Compras SC",
        "solicitacoes_compras",
        coluna_status="Status",
    )


@app.route("/governanca/backup-json")
@requer_login
def exportar_backup_json():
    dados = serializar_banco_completo()
    conteudo = json.dumps(dados, ensure_ascii=False, indent=2).encode("utf-8")
    memoria = BytesIO(conteudo)
    memoria.seek(0)
    nome = f"backup_completo_jamaica_{datetime.now():%Y%m%d_%H%M%S}.json"
    registrar_auditoria(
        "Governança",
        "Exportação Backup JSON",
        "Cópia completa do banco exportada",
    )
    return send_file(
        memoria,
        as_attachment=True,
        download_name=nome,
        mimetype="application/json",
    )


@app.route("/governanca/backup-excel-completo")
@requer_login
def exportar_backup_excel_completo():
    ordens = OrdemServico.query.order_by(OrdemServico.id.desc()).all()
    df_os = pd.DataFrame([
        {
            "Número O.S.": o.numero_os,
            "Solicitante": o.solicitante,
            "Máquina / Setor": o.setor or "",
            "Causa": o.titulo or "",
            "Prioridade": o.prioridade,
            "Responsável": o.responsavel or "",
            "Tipo Manutenção": o.tipo_manutencao or "",
            "Funcionários": o.funcionarios or "",
            "Tempo Apontado": (
                horas_decimais_para_hhmm(o.horas_trabalhadas)
                if o.horas_trabalhadas
                else ""
            ),
            "Laudo Técnico": o.observacao_tecnica or "",
            "Status": o.status,
            "Abertura": o.abertura,
            "Finalização": o.finalizacao or "",
            "Lead Time": o.tempo_total or "",
        }
        for o in ordens
    ])

    pecas = ItemEstoque.query.order_by(ItemEstoque.nome.asc()).all()
    df_est = pd.DataFrame([
        {
            "Código": p.codigo,
            "Peça": p.nome,
            "Categoria": p.categoria or "",
            "Localização": p.localizacao or "",
            "Saldo Atual": p.quantidade,
            "Estoque Mínimo": p.estoque_minimo,
            "Status": (
                "Baixo Estoque"
                if p.quantidade <= p.estoque_minimo
                else "Normal"
            ),
            "Última Atualização": p.ultima_atualizacao,
        }
        for p in pecas
    ])

    maquinas = MaquinaSpare.query.order_by(MaquinaSpare.maquina.asc()).all()
    rows_spare = []
    for m in maquinas:
        if m.pecas:
            for ps in m.pecas:
                rows_spare.append({
                    "Máquina": m.maquina,
                    "Modelo": m.modelo,
                    "Setor": m.setor or "",
                    "Peça": ps.nome_peca,
                    "Código": ps.codigo_peca or "",
                    "Qtd": ps.quantidade,
                    "Observação": ps.observacao or "",
                })
        else:
            rows_spare.append({
                "Máquina": m.maquina,
                "Modelo": m.modelo,
                "Setor": m.setor or "",
                "Peça": "",
                "Código": "",
                "Qtd": 0,
                "Observação": m.pecas_necessarias or "",
            })
    df_spare = pd.DataFrame(rows_spare)

    planos = PlanoPreventivo.query.all()
    hoje = date.today()
    df_pcm = pd.DataFrame([
        {
            "Máquina": p.maquina,
            "Tag": p.codigo_maquina or "",
            "Setor": p.setor or "",
            "Frequência": p.frequencia,
            "Próxima Execução": p.proxima_execucao,
            "Responsável": p.responsavel or "",
            "Status": status_plano(p, hoje),
            "Roteiro": p.atividade,
        }
        for p in planos
    ])

    compras = SolicitacaoCompra.query.order_by(
        SolicitacaoCompra.id.desc()
    ).all()
    rows_sc = []
    for c in compras:
        for idx, it in enumerate(c.itens, start=1):
            rows_sc.append({
                "Número S.C.": c.numero_sc,
                "Classificação": c.tipo_sc,
                "Status": c.status_aprovacao,
                "Data": c.data_solicitacao,
                "Solicitante": c.solicitante,
                "Centro de Custo": c.centro_custo,
                "Objeto": c.objeto_solicitacao,
                "Item": idx,
                "Descritivo": it.descritivo,
                "Qtd": it.quantidade,
                "Unidade": it.unidade,
                "Link": it.link_referencia or "",
            })
    df_sc = pd.DataFrame(rows_sc)

    registrar_auditoria(
        "Governança",
        "Exportação Backup Excel",
        "Planilha consolidada multi-abas gerada",
    )
    return gerar_excel_multiplas_abas(
        [
            ("Ordens de Servico", df_os, "Status"),
            ("Estoque", df_est, "Status"),
            ("Spare Parts", df_spare, None),
            ("Cronograma PCM", df_pcm, "Status"),
            ("Compras SC", df_sc, "Status"),
        ],
        "backup_geral_jamaica",
    )


@app.route("/governanca/backup")
@requer_login
def baixar_backup_db():
    if IS_POSTGRES or not os.path.exists(DB_PATH):
        flash(
            "O banco atual roda na nuvem (PostgreSQL). Utilize 'Exportar Banco"
            " Completo (.JSON)' ou 'Planilha Geral Multi-Abas (.XLSX)'."
        )
        return redirect(url_for("central_gestao", tab="governanca"))
    registrar_auditoria(
        "Governança",
        "Download Backup",
        "Cópia manual do banco os.db baixada",
    )
    return send_file(
        DB_PATH,
        as_attachment=True,
        download_name=f"os_backup_manual_{datetime.now():%Y%m%d_%H%M}.db",
    )


def obter_ip_rede():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


if __name__ == "__main__":
    porta = int(os.environ.get("PORT", 5000))
    ip_local = obter_ip_rede()
    print("=" * 68)
    print("  CENTRAL DE MANUTENÇÃO JAMAICA — INICIADA COM SUCESSO")
    print(f"  Banco de Dados : {'PostgreSQL (Nuvem)' if IS_POSTGRES else 'SQLite Local (os.db)'}")
    print(f"  Acesso Local   : http://127.0.0.1:{porta}")
    print(f"  Acesso na Rede : http://{ip_local}:{porta}")
    print("=" * 68)
    app.run(host="0.0.0.0", port=porta, debug=False, threaded=True)
