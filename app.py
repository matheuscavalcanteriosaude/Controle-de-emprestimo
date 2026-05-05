import sqlite3
import csv
import subprocess
import shutil
import os
import re
import io
import zipfile
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from io import StringIO
from datetime import datetime, timedelta, date
from docxtpl import DocxTemplate
from flask import (
    Flask, render_template, request, redirect, url_for, flash, Response, send_file, abort, make_response, jsonify, session
)

# - Constantes

UNIDADES_ROTA = [
    "CER Barra",
    "Complexo Municipal Rocha Faria",
    "CTI Pediátrico - Souza Aguiar",
    "CTI Pediátrico - Jesus",
    "UPA Costa Barros",
    "UPA Rocha Miranda",
    "UPA Madureira",
    "UPA Cidade de Deus",
    "UPA Engenho de Dentro",
    "UPA Del Castilho",
    "UPA Senador Camará",
    "UPA Vila Kennedy",
    "UPA Magalhães Bastos",
    "UPA Sepetiba",
    "UPA Paciência",
    "UPA João XXIII",
    "Hospital Municipal Ronaldo Gazolla",
    "Hospital Maternidade da Rocinha",
    "Hospital Federal do Andaraí",
    "Sede Administrativa da RioSaúde",
    "Outro",
]

APP_DB = "notebooks.db"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TERMOS_DIR = os.path.join(BASE_DIR, "termos")
TERMOS_GERADOS_DIR = os.path.join(BASE_DIR, "termos_gerados")
os.makedirs(TERMOS_GERADOS_DIR, exist_ok=True)

TIPO_EQUIPAMENTOS = [
    "Computador",
    "Monitor",
    "Televisão",
    "Impressora",
]

EMPRESAS_CHAMADO = [
    "Simpress",
    "Kaique",
    "Positivo",
    "Multi",
    "HP",
]

MARCAS_EQUIP = [
    "HP",
    "Positivo",
    "3Green",
    "Multi",
]
# Label de prioridades

def mapear_prioridade(prioridade_raw: str):
    """Recebe o texto vindo do Forms e devolve (label_curta, classe_css)."""
    if not prioridade_raw:
        return ("-", "badge-prioridade-baixa")

    p = prioridade_raw.lower()

    if "imediat" in p:   # "Imediata (Após 12h ...)"
        return ("Imediata", "badge-prioridade-imediata")
    if "alta" in p:
        return ("Alta", "badge-prioridade-alta")
    if "média" in p or "media" in p:
        return ("Média", "badge-prioridade-media")
    if "baixa" in p:
        return ("Baixa", "badge-prioridade-baixa")

    # fallback
    return (prioridade_raw, "badge-prioridade-baixa")


# - Gerador de termos

def find_soffice():
    candidatos = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    # se tiver no PATH
    if shutil.which("soffice"):
        return "soffice"
    for c in candidatos:
        if os.path.isfile(c):
            return c
    return None

SOFFICE_PATH = find_soffice()


app = Flask(__name__)
app.secret_key = "chave-super-secreta"  # troque se quiser

# ===================== CONTROLE DE ACESSO (DECORATORS) =====================

def check_password_change():
    """Função de apoio para verificar se a pessoa precisa mudar a senha"""
    return session.get("must_change_password") == 1

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("Por favor, faça login para acessar esta página.", "error")
            return redirect(url_for("login"))
        if check_password_change():
            flash("Por segurança, altere sua senha inicial antes de continuar.", "warning")
            return redirect(url_for("alterar_senha"))
        return f(*args, **kwargs)
    return decorated_function

def tecnico_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session: return redirect(url_for("login"))
        if check_password_change(): return redirect(url_for("alterar_senha"))
        
        if session.get("user_role") not in ["admin", "supervisor", "tecnico"]:
            flash("Acesso negado. Seu perfil não tem permissão para esta área.", "error")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session: return redirect(url_for("login"))
        if check_password_change(): return redirect(url_for("alterar_senha"))
        
        if session.get("user_role") not in ["admin", "supervisor"]:
            flash("Acesso negado. Requer privilégios de Administrador ou Supervisor.", "error")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return decorated_function

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads", "termos_fixos")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def safe_filename(name: str) -> str:
    # tira caracteres problemáticos de nome de arquivo
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    name = name.strip()
    if not name:
        name = "termo"
    return name

def gerar_termo_pdf(loan_id: int) -> str | None:
    """
    Usa o modelo .docx do termo para gerar um DOCX preenchido
    e, em seguida, converte para PDF usando LibreOffice.
    Retorna o caminho do PDF ou None se algo falhar.
    """
    # 1) gera o DOCX preenchido (já usa o seu modelo)
    docx_path = gerar_termo_docx(loan_id)
    if not docx_path or not os.path.isfile(docx_path):
        return None

    # 2) precisa do LibreOffice (soffice)
    global SOFFICE_PATH
    if not SOFFICE_PATH:
        return None  # depois tratamos isso na rota

    outdir = TERMOS_GERADOS_DIR
    cmd = [
        SOFFICE_PATH,
        "--headless",
        "--nologo",
        "--convert-to",
        "pdf",
        "--outdir",
        os.path.abspath(outdir),
        os.path.abspath(docx_path),
    ]
    try:
        subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        return None

    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    if os.path.isfile(pdf_path):
        return pdf_path
    return None


def gerar_termo_docx(loan_id: int) -> str | None:
    """
    Gera um termo DOCX a partir do modelo_termo.docx usando os dados do empréstimo.
    Retorna o caminho do arquivo gerado ou None se der algum problema.
    """
    template_path = os.path.join(TERMOS_DIR, "modelo_termo.docx")
    if not os.path.isfile(template_path):
        # se o modelo não existir, não faz nada
        return None

    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            l.id,
            l.colaborador_nome,
            l.colaborador_setor,
            l.data_hora_emprestimo,
            n.serial,
            n.modelo
        FROM loans l
        JOIN notebooks n ON n.id = l.notebook_id
        WHERE l.id = ?
        """,
        (loan_id,),
    )
    row = cur.fetchone()
    conn.close()

    if row is None:
        return None

    # prepara dados
    nome = row["colaborador_nome"]
    setor = row["colaborador_setor"] or ""
    serial = row["serial"] or ""
    modelo = row["modelo"] or ""

    try:
        dt = datetime.strptime(row["data_hora_emprestimo"], "%Y-%m-%d %H:%M:%S")
        data_emprestimo_br = dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        data_emprestimo_br = row["data_hora_emprestimo"]

    # contexto para o docxtpl → use os mesmos nomes de placeholders do seu modelo
    context = {
        "NOME": nome,
        "SETOR": setor,
        "SERIAL": serial,
        "MODELO": modelo,
        "DATA_EMPRESTIMO": data_emprestimo_br,
    }

    tpl = DocxTemplate(template_path)
    tpl.render(context)

    base = safe_filename(f"Termo_{loan_id}_{nome}")
    output_path = os.path.join(TERMOS_GERADOS_DIR, base + ".docx")
    tpl.save(output_path)

    return output_path


def get_connection():
    conn = sqlite3.connect(APP_DB)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cur = conn.cursor()

    # Tabela de rotas
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS rotas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_solicitacao TEXT NOT NULL,
            solicitante TEXT NOT NULL,
            unidade_origem TEXT NOT NULL,
            prioridade TEXT NOT NULL,
            destino TEXT NOT NULL,
            descricao_volume TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'pendente'
        )
        """
    )

    # Tabela de notebooks (estoque)
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS notebooks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            serial TEXT UNIQUE NOT NULL,
            modelo TEXT,
            status TEXT DEFAULT 'ativo'
        );
        """
    )

    # Tabela de empréstimos rotativos
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS loans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            notebook_id INTEGER NOT NULL,
            colaborador_nome TEXT NOT NULL,
            colaborador_setor TEXT,
            data_hora_emprestimo TEXT NOT NULL,
            data_hora_devolucao TEXT,
            status TEXT NOT NULL DEFAULT 'emprestado',
            FOREIGN KEY (notebook_id) REFERENCES notebooks (id)
        );
        """
    )

    # Tabela de agendamentos
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            notebook_id INTEGER NOT NULL,
            colaborador_nome TEXT NOT NULL,
            colaborador_setor TEXT,
            data_inicio TEXT NOT NULL,
            data_fim TEXT NOT NULL,
            inclui_som INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'agendado',
            FOREIGN KEY (notebook_id) REFERENCES notebooks (id)
        );
        """
    )

    # Tabela de Tickets (chamados)
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS tickets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_abertura TEXT NOT NULL,
            data_encerramento TEXT,
            status TEXT NOT NULL, -- 'aberto' ou 'fechado'

            tipo_equipamento TEXT,
            empresa TEXT,
            marca TEXT,
            modelo TEXT,
            numero_serie TEXT,
            defeito TEXT,

            endereco TEXT,
            cep TEXT,
            telefone TEXT,
            contatos TEXT,
            email TEXT,

            observacoes TEXT
        );
        """
    )

    # === NOVA TABELA: Empréstimos Fixos ===
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS fixed_loans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            notebook_id INTEGER NOT NULL,
            colaborador_nome TEXT NOT NULL,
            data_emprestimo TEXT NOT NULL,
            modem_4g INTEGER DEFAULT 0,
            modem_serial TEXT,
            termo_arquivo TEXT,
            status TEXT NOT NULL DEFAULT 'ativo',
            FOREIGN KEY (notebook_id) REFERENCES notebooks (id)
        );
        """
    )

# === NOVA TABELA: Usuários ===
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            nome_completo TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'operador',
            must_change_password INTEGER DEFAULT 1 -- <--- NOVO (1 = Sim, 0 = Não)
        );
        """
    )

    # Migração: Adiciona a coluna para quem já tem a tabela criada
    try:
        cur.execute("ALTER TABLE users ADD COLUMN must_change_password INTEGER DEFAULT 0;")
    except sqlite3.OperationalError:
        pass # A coluna já existe, segue o jogo

    # Verifica se existe algum usuário. Se não existir, cria o Admin padrão.
    cur.execute("SELECT COUNT(*) AS total FROM users")
    if cur.fetchone()["total"] == 0:
        # A senha do primeiro admin agora exigirá troca no primeiro login
        senha_criptografada = generate_password_hash("admin123")
        cur.execute(
            """
            INSERT INTO users (username, password_hash, nome_completo, role, must_change_password)
            VALUES (?, ?, ?, 'admin', 1)
            """,
            ("admin", senha_criptografada, "Administrador TI")
        )

    # Tenta adicionar a coluna setor (se der erro é porque ela já existe, então ignoramos)
    try:
        cur.execute("ALTER TABLE fixed_loans ADD COLUMN colaborador_setor TEXT;")
    except sqlite3.OperationalError:
        pass

    
    conn.commit()
    conn.close()


def calcular_tempo_limite(data_hora_emprestimo_str: str):
    """
    Calcula quanto tempo falta (ou passou) até o limite das 18h
    no dia do empréstimo.
    Retorna (texto_exibicao, nivel_cor).
    """
    try:
        dt_emprestimo = datetime.strptime(data_hora_emprestimo_str, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return ("-", "neutro")

    limite = dt_emprestimo.replace(hour=18, minute=0, second=0, microsecond=0)
    agora = datetime.now()
    diff_seg = (limite - agora).total_seconds()

    if diff_seg >= 0:
        horas = int(diff_seg // 3600)
        minutos = int((diff_seg % 3600) // 60)
        texto = f"Faltam {horas}h {minutos}min"

        if diff_seg > 4 * 3600:
            nivel = "verde"
        elif diff_seg > 1 * 3600:
            nivel = "amarelo"
        else:
            nivel = "vermelho"
    else:
        diff_seg = abs(diff_seg)
        horas = int(diff_seg // 3600)
        minutos = int((diff_seg % 3600) // 60)
        texto = f"Atrasado {horas}h {minutos}min"
        nivel = "atrasado"

    return texto, nivel

# ------------------------ ROTAS ----------------------------- #

# ===================== ROTAS DE AUTENTICAÇÃO =====================

@app.route("/login", methods=["GET", "POST"])
def login():
    # Se já estiver logado, manda pro dashboard
    if "user_id" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        username = request.form.get("username").strip()
        password = request.form.get("password")

        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE username = ?", (username,))
        user = cur.fetchone()
        conn.close()

        # Verifica se o usuário existe e se a senha bate com o hash salvo
        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["user_nome"] = user["nome_completo"]
            session["user_role"] = user["role"]
            session["must_change_password"] = user["must_change_password"] # <--- NOVO
            
            # Se precisar mudar, vai direto pra tela de alteração
            if session["must_change_password"] == 1:
                return redirect(url_for("alterar_senha"))
            
            flash(f"Bem-vindo(a), {user['nome_completo']}!", "success")
            return redirect(url_for("dashboard"))
        else:
            flash("Usuário ou senha incorretos.", "error")

    # Por padrão, vamos usar um template simples de login
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear() # Limpa tudo da sessão
    flash("Você saiu do sistema.", "success")
    return redirect(url_for("login"))

@app.route("/alterar-senha", methods=["GET", "POST"])
def alterar_senha():
    if "user_id" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        nova_senha = request.form.get("nova_senha")
        confirmar_senha = request.form.get("confirmar_senha")

        if nova_senha != confirmar_senha:
            flash("As senhas não coincidem. Tente novamente.", "error")
        elif len(nova_senha) < 6:
            flash("A senha deve ter no mínimo 6 caracteres.", "error")
        else:
            conn = get_connection()
            cur = conn.cursor()
            hashed_pw = generate_password_hash(nova_senha)
            # Atualiza a senha e diz que não precisa mais trocar
            cur.execute(
                "UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?",
                (hashed_pw, session["user_id"])
            )
            conn.commit()
            conn.close()

            session["must_change_password"] = 0
            flash("Senha alterada com sucesso!", "success")
            return redirect(url_for("dashboard"))

    return render_template("alterar_senha.html")

# Rotas DOP

#@app.route("/rotas")#
def lista_rotas():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            id,
            data_solicitacao,
            solicitante,
            unidade_origem,
            prioridade,
            destino,
            descricao_volume,
            status
        FROM rotas
        ORDER BY data_solicitacao DESC
        """
    )
    rows = cur.fetchall()
    conn.close()

    rotas = []
    for r in rows:
        d = dict(r)
        dt = datetime.strptime(d["data_solicitacao"], "%Y-%m-%d %H:%M:%S")
        d["data_solicitacao_br"] = dt.strftime("%d/%m/%Y %H:%M")

        label, css = mapear_prioridade(d.get("prioridade", ""))
        d["prioridade_label"] = label
        d["prioridade_css"] = css

        rotas.append(d)

    return render_template("rotas.html", rotas=rotas)


# Rotas DOP - Email

@app.route("/api/rotas_email", methods=["POST"])
def api_rotas_email():
    """
    Endpoint chamado pelo Apps Script.
    Recebe JSON com:
      - solicitante
      - unidade_origem
      - prioridade
      - destino
      - descricao_volume
    Grava na tabela rotas.
    """
    data = request.get_json(force=True)

    obrigatorios = [
        "solicitante",
        "unidade_origem",
        "prioridade",
        "destino",
        "descricao_volume",
    ]

    for campo in obrigatorios:
        if not data.get(campo):
            return jsonify(
                {"ok": False, "erro": f"Campo '{campo}' obrigatório"}
            ), 400

    # normaliza textos (tira espaços a mais)
    solicitante = data["solicitante"].strip()
    unidade_origem = data["unidade_origem"].strip()
    prioridade = data["prioridade"].strip()
    destino = data["destino"].strip()
    descricao_volume = data["descricao_volume"].strip()

    conn = get_connection()
    cur = conn.cursor()

    # data/hora que a rota entrou no sistema
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cur.execute(
        """
        INSERT INTO rotas (
            data_solicitacao,
            solicitante,
            unidade_origem,
            prioridade,
            destino,
            descricao_volume,
            status
        )
        VALUES (?, ?, ?, ?, ?, ?, 'pendente')
        """,
        (
            agora,
            solicitante,
            unidade_origem,
            prioridade,
            destino,
            descricao_volume,
        ),
    )

    conn.commit()
    conn.close()

    return jsonify({"ok": True})

# Chamados

@app.route("/chamados/novo", methods=["GET", "POST"])
def novo_chamado():
    if request.method == "POST":
        form = request.form

        data_abertura = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        tipo_equipamento = form.get("tipo_equipamento", "").strip()
        empresa = form.get("empresa", "").strip()
        marca = form.get("marca", "").strip()
        modelo = form.get("modelo", "").strip()
        numero_serie = form.get("numero_serie", "").strip()
        defeito = form.get("defeito", "").strip()

        endereco = form.get("endereco", "").strip()
        cep = form.get("cep", "").strip()
        telefone = form.get("telefone", "").strip()
        contatos = form.get("contatos", "").strip()
        email = form.get("email", "").strip()
        observacoes = form.get("observacoes", "").strip()

        # validação mínima
        if not tipo_equipamento or not numero_serie or not defeito:
            flash("Informe tipo de equipamento, número de série e defeito apresentado.", "error")
            return render_template(
                "chamado_form.html",
                tipos=TIPO_EQUIPAMENTOS,
                empresas=EMPRESAS_CHAMADO,
                marcas=MARCAS_EQUIP,
            )

        conn = get_connection()
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO tickets (
                data_abertura,
                status,
                tipo_equipamento,
                empresa,
                marca,
                modelo,
                numero_serie,
                defeito,
                endereco,
                cep,
                telefone,
                contatos,
                email,
                observacoes
            )
            VALUES (?, 'aberto', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                data_abertura,
                tipo_equipamento,
                empresa,
                marca,
                modelo,
                numero_serie,
                defeito,
                endereco,
                cep,
                telefone,
                contatos,
                email,
                observacoes,
            ),
        )
        conn.commit()
        conn.close()

        flash("Chamado aberto com sucesso.", "success")
        return redirect(url_for("lista_chamados"))

    # GET
    return render_template(
        "chamado_form.html",
        tipos=TIPO_EQUIPAMENTOS,
        empresas=EMPRESAS_CHAMADO,
        marcas=MARCAS_EQUIP,
    )

@app.route("/chamados")
def lista_chamados():
    status = request.args.get("status", "aberto")  # padrão: abertos

    conn = get_connection()
    cur = conn.cursor()
    if status == "todos":
        cur.execute("SELECT * FROM tickets ORDER BY data_abertura DESC")
    else:
        cur.execute(
            "SELECT * FROM tickets WHERE status = ? ORDER BY data_abertura DESC",
            (status,),
        )
    rows = cur.fetchall()
    conn.close()

    hoje = date.today()
    chamados = []
    for r in rows:
        d = dict(r)
        dt_ab = datetime.strptime(d["data_abertura"], "%Y-%m-%d %H:%M:%S")
        d["data_abertura_br"] = dt_ab.strftime("%d/%m/%Y %H:%M")
        d["dias_aberto"] = (hoje - dt_ab.date()).days
        chamados.append(d)

    return render_template(
        "chamados.html",
        chamados=chamados,
        filtro_status=status,
    )

@app.route("/chamados/fechar/<int:ticket_id>", methods=["POST"])
def fechar_chamado(ticket_id):
    conn = get_connection()
    cur = conn.cursor()
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        UPDATE tickets
        SET status = 'fechado', data_encerramento = ?
        WHERE id = ?
        """,
        (agora, ticket_id),
    )
    conn.commit()
    conn.close()

    flash("Chamado encerrado com sucesso.", "success")
    return redirect(url_for("lista_chamados"))


@app.route("/termo/<int:loan_id>")
def termo_opcoes(loan_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            l.id,
            l.colaborador_nome,
            l.colaborador_setor,
            l.data_hora_emprestimo,
            n.serial,
            n.modelo
        FROM loans l
        JOIN notebooks n ON n.id = l.notebook_id
        WHERE l.id = ?
        """,
        (loan_id,),
    )
    row = cur.fetchone()
    conn.close()

    if row is None:
        abort(404)

    try:
        dt = datetime.strptime(row["data_hora_emprestimo"], "%Y-%m-%d %H:%M:%S")
        data_br = dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        data_br = row["data_hora_emprestimo"]

    return render_template(
        "termo.html",
        emprestimo=row,
        data_br=data_br,
    )


@app.route("/termo/<int:loan_id>/download")
def baixar_termo(loan_id):
    caminho = gerar_termo_docx(loan_id)
    if not caminho or not os.path.isfile(caminho):
        flash("Não foi possível gerar o termo para este empréstimo.", "error")
        return redirect(url_for("dashboard"))

    return send_file(
        caminho,
        as_attachment=True,
        download_name=os.path.basename(caminho),
    )


@app.route("/")
def index():
    # Se o usuário já estiver logado, joga ele pro Dashboard
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    
    # Se não estiver logado, joga ele para a tela de Login naturalmente
    return redirect(url_for("login"))

@app.route("/termo/<int:loan_id>/visualizar")
def visualizar_termo(loan_id):
    pdf_path = gerar_termo_pdf(loan_id)
    if not pdf_path or not os.path.isfile(pdf_path):
        flash(
            "Não foi possível gerar o PDF do termo. "
            "Verifique se o LibreOffice está instalado ou use a opção de download em DOCX.",
            "error",
        )
        return redirect(url_for("termo_opcoes", loan_id=loan_id))

    # abre o PDF diretamente no navegador (sem forçar download)
    return send_file(pdf_path, mimetype="application/pdf")

# Rotas - manual

@app.route("/rotas", methods=["GET", "POST"])
def lista_rotas():
    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":
        solicitante = request.form.get("solicitante", "").strip()
        unidade_origem = request.form.get("unidade_origem", "").strip()
        prioridade = request.form.get("prioridade", "").strip()
        destino = request.form.get("destino", "").strip()
        descricao_volume = request.form.get("descricao_volume", "").strip()

        if not (solicitante and unidade_origem and prioridade and destino and descricao_volume):
            flash("Preencha todos os campos obrigatórios da rota.", "error")
        else:
            data_solicitacao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cur.execute(
                """
                INSERT INTO rotas (
                    data_solicitacao,
                    solicitante,
                    unidade_origem,
                    prioridade,
                    destino,
                    descricao_volume,
                    status
                )
                VALUES (?, ?, ?, ?, ?, ?, 'pendente')
                """,
                (
                    data_solicitacao,
                    solicitante,
                    unidade_origem,
                    prioridade,
                    destino,
                    descricao_volume,
                ),
            )
            conn.commit()
            flash("Rota cadastrada com sucesso.", "success")

        return redirect(url_for("lista_rotas"))

    # GET -> lista rotas
    cur.execute(
        """
        SELECT
            id,
            data_solicitacao,
            solicitante,
            unidade_origem,
            prioridade,
            destino,
            descricao_volume,
            status
        FROM rotas
        ORDER BY data_solicitacao DESC
        """
    )
    rows = cur.fetchall()
    conn.close()

    rotas = []
    for r in rows:
        d = dict(r)
        dt = datetime.strptime(d["data_solicitacao"], "%Y-%m-%d %H:%M:%S")
        d["data_solicitacao_br"] = dt.strftime("%d/%m/%Y %H:%M")

        label, css = mapear_prioridade(d.get("prioridade", ""))
        d["prioridade_label"] = label
        d["prioridade_css"] = css

        rotas.append(d)

    return render_template(
        "rotas.html",
        rotas=rotas,
        unidades=UNIDADES_ROTA,
    )



@app.route("/rotas/enviar/<int:rota_id>", methods=["POST"])
def marcar_rota_enviada(rota_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "UPDATE rotas SET status = 'enviada' WHERE id = ?",
        (rota_id,),
    )
    conn.commit()
    conn.close()
    flash("Rota marcada como enviada.", "success")
    return redirect(url_for("lista_rotas"))



# ===================== CADASTRO / LISTA DE NOTEBOOKS =====================

@app.route("/notebooks/novo", methods=["GET", "POST"])
@tecnico_required
def novo_notebook():
    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":
        serial = request.form.get("serial", "").strip()
        modelo = request.form.get("modelo", "").strip()
        tipo_cadastro = request.form.get("tipo_cadastro") # 'rotativo' ou 'fixo'

        if not serial:
            flash("O número de série é obrigatório.", "error")
            conn.close()
            return redirect(url_for("novo_notebook"))

        # Verifica se o serial já existe no banco
        cur.execute("SELECT id, status FROM notebooks WHERE serial = ?", (serial,))
        existente = cur.fetchone()
        notebook_id = None

        if existente:
            if existente["status"] == "inativo":
                # Reativa o notebook
                cur.execute("UPDATE notebooks SET status = 'ativo', modelo = ? WHERE id = ?", (modelo, existente["id"]))
                notebook_id = existente["id"]
                flash("Notebook reativado com sucesso!", "success")
            else:
                flash("Já existe um notebook na base de ativos ou fixos com esse número de série.", "error")
                conn.close()
                return redirect(url_for("novo_notebook"))
        else:
            # Novo cadastro
            cur.execute("INSERT INTO notebooks (serial, modelo, status) VALUES (?, ?, 'ativo')", (serial, modelo))
            notebook_id = cur.lastrowid
            flash("Notebook cadastrado com sucesso!", "success")

# Se o usuário escolheu "Fixo", já registra o empréstimo fixo e anexa o arquivo
        if tipo_cadastro == "fixo" and notebook_id:
            colaborador_nome = request.form.get("colaborador_nome", "").strip()
            colaborador_setor = request.form.get("colaborador_setor", "").strip() # <--- NOVO
            data_emprestimo = request.form.get("data_emprestimo", "").strip()
            modem_4g = 1 if request.form.get("modem_4g") == "on" else 0
            modem_serial = request.form.get("modem_serial", "").strip() if modem_4g else ""

            arquivo = request.files.get("termo_arquivo")
            nome_arquivo_salvo = ""
            
            if arquivo and arquivo.filename != "":
                nome_original = secure_filename(arquivo.filename)
                nome_arquivo_salvo = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{nome_original}"
                caminho_salvo = os.path.join(app.config['UPLOAD_FOLDER'], nome_arquivo_salvo)
                arquivo.save(caminho_salvo)

            # Grava na tabela de empréstimos fixos (AGORA COM SETOR)
            cur.execute(
                """
                INSERT INTO fixed_loans (
                    notebook_id, colaborador_nome, colaborador_setor, data_emprestimo, 
                    modem_4g, modem_serial, termo_arquivo, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 'ativo')
                """,
                (notebook_id, colaborador_nome, colaborador_setor, data_emprestimo, modem_4g, modem_serial, nome_arquivo_salvo)
            )
            # Muda o status do equipamento para 'fixo'
            cur.execute("UPDATE notebooks SET status = 'fixo' WHERE id = ?", (notebook_id,))
            flash(f"Empréstimo fixo para {colaborador_nome} ({colaborador_setor}) registrado!", "success")

        conn.commit()
        conn.close()
        return redirect(url_for("novo_notebook"))

    # ==================== GET (Renderiza as Duas Listas) ====================
    # 1. Rotativos (apenas ativos e inativos)
    cur.execute(
        "SELECT id, serial, modelo, status FROM notebooks WHERE status IN ('ativo', 'inativo') ORDER BY status DESC, serial"
    )
    rotativos = cur.fetchall()

# 2. Fixos (Agora puxando o f.colaborador_setor)
    cur.execute(
        """
        SELECT n.id as notebook_id, n.serial, n.modelo, 
               f.id as fixo_id, f.colaborador_nome, f.colaborador_setor, f.termo_arquivo 
        FROM notebooks n
        JOIN fixed_loans f ON n.id = f.notebook_id
        WHERE n.status = 'fixo' AND f.status = 'ativo'
        ORDER BY n.serial
        """
    )
    fixos = cur.fetchall()
    conn.close()

    return render_template("notebook_form.html", rotativos=rotativos, fixos=fixos)

# Rota para baixar o anexo (Adicione se ainda não tiver)
@app.route("/emprestimos-fixos/termo/<filename>")
def baixar_termo_fixo(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

# Rota para devolver o Fixo e retorná-lo para os Rotativos
@app.route("/emprestimos-fixos/devolver/<int:fixo_id>", methods=["POST"])
def devolver_fixo(fixo_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT notebook_id FROM fixed_loans WHERE id = ?", (fixo_id,))
    row = cur.fetchone()
    
    if row:
        notebook_id = row["notebook_id"]
        cur.execute("UPDATE fixed_loans SET status = 'devolvido' WHERE id = ?", (fixo_id,))
        cur.execute("UPDATE notebooks SET status = 'ativo' WHERE id = ?", (notebook_id,))
        conn.commit()
        flash("Equipamento devolvido! Agora ele é um notebook rotativo disponível no estoque.", "success")
        
    conn.close()
    return redirect(url_for("novo_notebook"))

# ===================== AÇÕES ROTATIVOS =====================

@app.route("/notebooks/remover/<int:notebook_id>", methods=["POST"])
@tecnico_required
def remover_notebook(notebook_id):
    conn = get_connection()
    cur = conn.cursor()

    # Verifica se há empréstimo ativo antes de inativar
    cur.execute(
        "SELECT COUNT(*) AS total FROM loans WHERE notebook_id = ? AND status = 'emprestado'",
        (notebook_id,),
    )
    em_uso = cur.fetchone()["total"]

    if em_uso > 0:
        flash("Não é possível inativar: notebook está com empréstimo rotativo ativo.", "error")
    else:
        cur.execute("UPDATE notebooks SET status = 'inativo' WHERE id = ?", (notebook_id,))
        conn.commit()
        flash("Notebook inativado com sucesso.", "success")
        
    conn.close()
    return redirect(url_for("novo_notebook"))


@app.route("/notebooks/reativar/<int:notebook_id>", methods=["POST"])
@tecnico_required
def reativar_notebook(notebook_id):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("UPDATE notebooks SET status = 'ativo' WHERE id = ?", (notebook_id,))
    conn.commit()
    conn.close()

    flash("Notebook reativado com sucesso.", "success")
    return redirect(url_for("novo_notebook"))


# ===================== EMPRÉSTIMO =====================

@app.route("/emprestimo", methods=["GET", "POST"])
def emprestimo():
    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":
        notebook_id = request.form.get("notebook_id")
        colaborador_nome = request.form.get("colaborador_nome", "").strip()
        colaborador_setor = request.form.get("colaborador_setor", "").strip()

        if not notebook_id or not colaborador_nome:
            flash("Selecione o notebook e informe o nome do colaborador.", "error")
            conn.close()
            return redirect(url_for("emprestimo"))

        # Verifica se notebook já está emprestado
        cur.execute(
            "SELECT COUNT(*) AS total FROM loans WHERE notebook_id = ? AND status = 'emprestado'",
            (notebook_id,),
        )
        ja_emprestado = cur.fetchone()["total"]

        if ja_emprestado > 0:
            flash(
                "Este notebook já está emprestado. Não é possível emprestar novamente.",
                "error",
            )
            conn.close()
            return redirect(url_for("emprestimo"))

        data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cur.execute(
            """
            INSERT INTO loans (
                notebook_id, colaborador_nome, colaborador_setor,
                data_hora_emprestimo, status
            )
            VALUES (?, ?, ?, ?, 'emprestado')
            """,
            (notebook_id, colaborador_nome, colaborador_setor, data_hora),
        )

        loan_id = cur.lastrowid

        conn.commit()
        conn.close()

        flash("Empréstimo registrado com sucesso.", "success")
        return redirect(url_for("termo_opcoes", loan_id=loan_id))



    # GET → notebooks ativos e não emprestados
    cur.execute(
        """
        SELECT id, serial, modelo
        FROM notebooks
        WHERE status = 'ativo'
        AND id NOT IN (
            SELECT notebook_id FROM loans WHERE status = 'emprestado'
        )
        ORDER BY serial
        """
    )
    notebooks = cur.fetchall()
    conn.close()

    return render_template("loan_form.html", notebooks=notebooks)


# ===================== DEVOLUÇÃO =====================

@app.route("/devolver/<int:loan_id>")
def devolver(loan_id):
    conn = get_connection()
    cur = conn.cursor()

    data_hora_dev = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cur.execute(
        """
        UPDATE loans
        SET status = 'devolvido',
            data_hora_devolucao = ?
        WHERE id = ? AND status = 'emprestado'
        """,
        (data_hora_dev, loan_id),
    )
    conn.commit()
    conn.close()

    flash("Devolução registrada.", "success")
    return redirect(url_for("dashboard"))


# ===================== AGENDAMENTOS =====================

@app.route("/agendamentos/emprestar/<int:schedule_id>", methods=["GET", "POST"])
def emprestar_agendamento(schedule_id):
    conn = get_connection()
    cur = conn.cursor()

    # Busca o agendamento
    cur.execute(
        """
        SELECT
            id,
            notebook_id,
            colaborador_nome,
            colaborador_setor,
            data_inicio,
            data_fim,
            inclui_som,
            status
        FROM schedules
        WHERE id = ?
        """,
        (schedule_id,),
    )
    sched = cur.fetchone()

    if sched is None or sched["status"] != "agendado":
        conn.close()
        flash("Agendamento não encontrado ou já utilizado/cancelado.", "error")
        return redirect(url_for("agendamentos"))

    # Datas bonitinhas (já usadas no GET)
    dt_i = datetime.strptime(sched["data_inicio"], "%Y-%m-%d %H:%M:%S")
    dt_f = datetime.strptime(sched["data_fim"], "%Y-%m-%d %H:%M:%S")
    agendamento = dict(sched)
    agendamento["data_inicio_br"] = dt_i.strftime("%d/%m/%Y %H:%M")
    agendamento["data_fim_br"] = dt_f.strftime("%d/%m/%Y %H:%M")

    if request.method == "POST":
        notebook_id = request.form.get("notebook_id")

        if not notebook_id:
            conn.close()
            flash("Selecione um notebook para efetivar o empréstimo.", "error")
            return redirect(url_for("emprestar_agendamento", schedule_id=schedule_id))

        # Confere se o notebook existe e está ativo
        cur.execute(
            "SELECT id, status FROM notebooks WHERE id = ?",
            (notebook_id,),
        )
        nb = cur.fetchone()
        if nb is None or nb["status"] != "ativo":
            conn.close()
            flash("Notebook selecionado não está disponível.", "error")
            return redirect(url_for("emprestar_agendamento", schedule_id=schedule_id))

        # Confere se não está emprestado
        cur.execute(
            "SELECT COUNT(*) AS total FROM loans WHERE notebook_id = ? AND status = 'emprestado'",
            (notebook_id,),
        )
        em_uso = cur.fetchone()["total"]
        if em_uso > 0:
            conn.close()
            flash("Notebook selecionado já está emprestado no momento.", "error")
            return redirect(url_for("emprestar_agendamento", schedule_id=schedule_id))

        # Confere se não tem conflito de agendamento nesse horário
        cur.execute(
            """
            SELECT COUNT(*) AS total
            FROM schedules
            WHERE notebook_id = ?
              AND status IN ('agendado', 'em_uso')
              AND id <> ?
              AND NOT (data_fim <= ? OR data_inicio >= ?)
            """,
            (
                notebook_id,
                schedule_id,
                sched["data_inicio"],
                sched["data_fim"],
            ),
        )
        conflito = cur.fetchone()["total"]
        if conflito > 0:
            conn.close()
            flash("Notebook selecionado possui outro agendamento nesse horário.", "error")
            return redirect(url_for("emprestar_agendamento", schedule_id=schedule_id))

        # ---- Cria o empréstimo de fato ----
        data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur.execute(
            """
            INSERT INTO loans (
                notebook_id,
                colaborador_nome,
                colaborador_setor,
                data_hora_emprestimo,
                status
            )
            VALUES (?, ?, ?, ?, 'emprestado')
            """,
            (
                notebook_id,
                sched["colaborador_nome"],
                sched["colaborador_setor"],
                data_hora,
            ),
        )
        loan_id = cur.lastrowid  # <- ID do empréstimo criado

        # Atualiza o agendamento para 'em_uso' e guarda o notebook escolhido
        cur.execute(
            "UPDATE schedules SET status = 'em_uso', notebook_id = ? WHERE id = ?",
            (notebook_id, schedule_id),
        )

        conn.commit()
        conn.close()

        flash("Empréstimo efetuado a partir do agendamento.", "success")
        # Agora vai para a mesma tela de termo do fluxo normal
        return redirect(url_for("termo_opcoes", loan_id=loan_id))

    # ============ MÉTODO GET: monta lista de notebooks disponíveis ============

    # (resto do código GET que monta notebooks_disponiveis e renderiza
    #  efeti_var_agendamento.html permanece igual)
    cur.execute(
        "SELECT id, serial, modelo FROM notebooks WHERE status = 'ativo' ORDER BY serial"
    )
    notebooks = cur.fetchall()
    notebooks_disponiveis = []
    for nb in notebooks:
        cur.execute(
            "SELECT COUNT(*) AS total FROM loans WHERE notebook_id = ? AND status = 'emprestado'",
            (nb["id"],),
        )
        em_uso = cur.fetchone()["total"]
        if em_uso > 0:
            continue

        cur.execute(
            """
            SELECT COUNT(*) AS total
            FROM schedules
            WHERE notebook_id = ?
              AND status IN ('agendado', 'em_uso')
              AND id <> ?
              AND NOT (data_fim <= ? OR data_inicio >= ?)
            """,
            (
                nb["id"],
                schedule_id,
                sched["data_inicio"],
                sched["data_fim"],
            ),
        )
        conflito = cur.fetchone()["total"]
        if conflito > 0:
            continue

        notebooks_disponiveis.append(nb)

    conn.close()

    if not notebooks_disponiveis:
        flash("Nenhum notebook disponível para esse intervalo no momento.", "error")
        return redirect(url_for("agendamentos"))

    return render_template(
        "efetivar_agendamento.html",
        agendamento=agendamento,
        notebooks_disponiveis=notebooks_disponiveis,
    )




@app.route("/agendamentos", methods=["GET", "POST"])
def agendamentos():
    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":
        # NÃO pegamos mais notebook_id do formulário
        colaborador_nome = request.form.get("colaborador_nome", "").strip()
        colaborador_setor = request.form.get("colaborador_setor", "").strip()
        inicio_str = request.form.get("data_inicio", "").strip()
        fim_str = request.form.get("data_fim", "").strip()
        inclui_som = 1 if request.form.get("inclui_som") == "on" else 0

        if not (colaborador_nome and inicio_str and fim_str):
            flash("Preencha o nome do colaborador e os horários de início e fim.", "error")
            conn.close()
            return redirect(url_for("agendamentos"))

        try:
            dt_inicio = datetime.fromisoformat(inicio_str)  # 2025-11-24T13:30
            dt_fim = datetime.fromisoformat(fim_str)
        except ValueError:
            flash("Formato de data/hora inválido.", "error")
            conn.close()
            return redirect(url_for("agendamentos"))

        if dt_fim <= dt_inicio:
            flash("Horário final deve ser maior que o horário inicial.", "error")
            conn.close()
            return redirect(url_for("agendamentos"))

        inicio_db = dt_inicio.strftime("%Y-%m-%d %H:%M:%S")
        fim_db = dt_fim.strftime("%Y-%m-%d %H:%M:%S")

        # 🔹 1) pega todos os notebooks ativos
        cur.execute(
            "SELECT id, serial, modelo FROM notebooks WHERE status = 'ativo' ORDER BY id"
        )
        notebooks = cur.fetchall()

        if not notebooks:
            flash("Não há notebooks ativos cadastrados para agendar.", "error")
            conn.close()
            return redirect(url_for("agendamentos"))

        # 🔹 2) tenta encontrar um notebook livre no intervalo
        notebook_escolhido_id = None

        for nb in notebooks:
            cur.execute(
                """
                SELECT COUNT(*) AS total
                FROM schedules
                WHERE notebook_id = ?
                  AND status = 'agendado'
                  AND NOT (data_fim <= ? OR data_inicio >= ?)
                """,
                (nb["id"], inicio_db, fim_db),
            )
            conflito = cur.fetchone()["total"]
            if conflito == 0:
                notebook_escolhido_id = nb["id"]
                break

        # 🔹 3) se nenhum notebook estiver livre, bloqueia o agendamento
        if notebook_escolhido_id is None:
            flash(
                "Não há notebooks disponíveis nesse intervalo. Todos já estão agendados.",
                "error",
            )
            conn.close()
            return redirect(url_for("agendamentos"))

        # 🔹 4) grava o agendamento usando o notebook escolhido automaticamente
        cur.execute(
            """
            INSERT INTO schedules (
                notebook_id, colaborador_nome, colaborador_setor,
                data_inicio, data_fim, inclui_som, status
            )
            VALUES (?, ?, ?, ?, ?, ?, 'agendado')
            """,
            (
                notebook_escolhido_id,
                colaborador_nome,
                colaborador_setor,
                inicio_db,
                fim_db,
                inclui_som,
            ),
        )
        conn.commit()
        conn.close()

        flash("Agendamento criado com sucesso!", "success")
        return redirect(url_for("agendamentos"))
    
    
    # ================== BLOCO GET (pode manter o seu, mas já deixo alinhado) ==================

    # notebooks ativos só pra mostrar capacidade (opcional)
    cur.execute("SELECT COUNT(*) AS total FROM notebooks WHERE status = 'ativo'")
    total_notebooks = cur.fetchone()["total"]

    # Próximos agendamentos
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        SELECT
            s.id,
            n.serial,
            n.modelo,
            s.colaborador_nome,
            s.colaborador_setor,
            s.data_inicio,
            s.data_fim,
            s.inclui_som,
            s.status
        FROM schedules s
        JOIN notebooks n ON n.id = s.notebook_id
        WHERE s.status = 'agendado'
          AND s.data_fim >= ?
        ORDER BY s.data_inicio
        """,
        (agora,),
    )
    rows = cur.fetchall()
    conn.close()

    proximos = []
    hoje = datetime.now().date()
    for r in rows:
        d = dict(r)
        dt_i = datetime.strptime(d["data_inicio"], "%Y-%m-%d %H:%M:%S")
        dt_f = datetime.strptime(d["data_fim"], "%Y-%m-%d %H:%M:%S")
        d["data_inicio_br"] = dt_i.strftime("%d/%m/%Y %H:%M")
        d["data_fim_br"] = dt_f.strftime("%d/%m/%Y %H:%M")
        d["eh_hoje"] = (dt_i.date() == hoje)
        proximos.append(d)

    return render_template(
        "agendamentos.html",
        total_notebooks=total_notebooks,
        proximos_agendamentos=proximos,
    )


@app.route("/agendamentos/cancelar/<int:schedule_id>")
def cancelar_agendamento(schedule_id):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        UPDATE schedules
        SET status = 'cancelado'
        WHERE id = ? AND status = 'agendado'
        """,
        (schedule_id,),
    )
    conn.commit()
    conn.close()

    flash("Agendamento cancelado.", "success")
    return redirect(url_for("agendamentos"))


# ===================== DASHBOARD =====================

@app.route("/dashboard")
def dashboard():
    conn = get_connection()
    cur = conn.cursor()

    # --- Total de notebooks ativos (estoque) ---
    cur.execute("SELECT COUNT(*) AS total FROM notebooks WHERE status = 'ativo'")
    total_notebooks = cur.fetchone()["total"]

    # --- Empréstimos ativos ---
    cur.execute(
        """
        SELECT
            l.id,
            n.serial,
            n.modelo,
            l.colaborador_nome,
            l.colaborador_setor,
            l.data_hora_emprestimo
        FROM loans l
        JOIN notebooks n ON n.id = l.notebook_id
        WHERE l.status = 'emprestado'
        ORDER BY l.data_hora_emprestimo DESC
        """
    )
    rows_ativos = cur.fetchall()

    emprestimos_ativos = []
    for r in rows_ativos:
        d = dict(r)
        texto_limite, nivel_limite = calcular_tempo_limite(d["data_hora_emprestimo"])
        d["tempo_limite"] = texto_limite
        d["nivel_limite"] = nivel_limite
        emprestimos_ativos.append(d)

    total_emprestados = len(emprestimos_ativos)
    disponiveis = total_notebooks - total_emprestados

    # --- Histórico recente (últimos 10) ---
    cur.execute(
        """
        SELECT
            l.id,
            n.serial,
            n.modelo,
            l.colaborador_nome,
            l.colaborador_setor,
            l.data_hora_emprestimo,
            l.data_hora_devolucao,
            l.status
        FROM loans l
        JOIN notebooks n ON n.id = l.notebook_id
        ORDER BY l.id DESC
        LIMIT 10
        """
    )
    historico = cur.fetchall()

    # --- Próximos agendamentos (para o dashboard) ---
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        SELECT
            s.id,
            n.serial,
            n.modelo,
            s.colaborador_nome,
            s.colaborador_setor,
            s.data_inicio,
            s.data_fim,
            s.inclui_som
        FROM schedules s
        JOIN notebooks n ON n.id = s.notebook_id
        WHERE s.status = 'agendado'
          AND s.data_fim >= ?
        ORDER BY s.data_inicio
        LIMIT 10
        """,
        (agora,),
    )
    rows_sched = cur.fetchall()

    # --- Chamados em aberto (KPI + lista resumida) ---
    cur.execute(
        """
        SELECT COUNT(*) AS total
        FROM tickets
        WHERE status = 'aberto'
        """
    )
    total_chamados_abertos = cur.fetchone()["total"]

    cur.execute(
        """
        SELECT
            id,
            tipo_equipamento,
            numero_serie,
            empresa,
            data_abertura
        FROM tickets
        WHERE status = 'aberto'
        ORDER BY data_abertura DESC
        LIMIT 5
        """
    )
    rows_chamados = cur.fetchall()

    # --- Rotas recentes para o card "Rotas solicitadas" ---
    cur.execute(
        """
        SELECT
            id,
            data_solicitacao,
            solicitante,
            unidade_origem,
            prioridade,
            destino,
            descricao_volume,
            status
        FROM rotas
        WHERE status = 'pendente'
        ORDER BY data_solicitacao DESC
        LIMIT 10
        """
    )
    rows_rotas = cur.fetchall()

    conn.close()

    hoje = datetime.now().date()

    # --- Monta proximos_agendamentos ---
    proximos_agendamentos = []
    for r in rows_sched:
        d = dict(r)
        dt_i = datetime.strptime(d["data_inicio"], "%Y-%m-%d %H:%M:%S")
        dt_f = datetime.strptime(d["data_fim"], "%Y-%m-%d %H:%M:%S")
        d["data_inicio_br"] = dt_i.strftime("%d/%m/%Y %H:%M")
        d["data_fim_br"] = dt_f.strftime("%d/%m/%Y %H:%M")
        d["eh_hoje"] = (dt_i.date() == hoje)
        proximos_agendamentos.append(d)

    # --- Monta chamados_abertos ---
    chamados_abertos = []
    for r in rows_chamados:
        d = dict(r)
        dt = datetime.strptime(d["data_abertura"], "%Y-%m-%d %H:%M:%S")
        d["data_abertura_br"] = dt.strftime("%d/%m/%Y %H:%M")
        d["dias_aberto"] = (hoje - dt.date()).days
        chamados_abertos.append(d)

    # --- Monta rotas_dashboard ---
    rotas_dashboard = []
    for r in rows_rotas:
        d = dict(r)
        try:
            dt = datetime.strptime(d["data_solicitacao"], "%Y-%m-%d %H:%M:%S")
            d["data_solicitacao_br"] = dt.strftime("%d/%m/%Y %H:%M")
        except Exception:
            d["data_solicitacao_br"] = d["data_solicitacao"]

        label, css = mapear_prioridade(d.get("prioridade", ""))
        d["prioridade_label"] = label
        d["prioridade_css"] = css

        rotas_dashboard.append(d)

    return render_template(
        "dashboard.html",
        total_notebooks=total_notebooks,
        total_emprestados=total_emprestados,
        disponiveis=disponiveis,
        emprestimos_ativos=emprestimos_ativos,
        historico=historico,
        proximos_agendamentos=proximos_agendamentos,
        total_chamados_abertos=total_chamados_abertos,
        chamados_abertos=chamados_abertos,
        rotas_dashboard=rotas_dashboard,
    )


@app.route("/tv")
def tv_dashboard():
    conn = get_connection()
    cur = conn.cursor()

    # Total de notebooks
    cur.execute("SELECT COUNT(*) AS total FROM notebooks WHERE status = 'ativo'")
    total_notebooks = cur.fetchone()["total"]

    # Empréstimos ativos
    cur.execute(
        """
        SELECT
            l.id,
            n.serial,
            n.modelo,
            l.colaborador_nome,
            l.colaborador_setor,
            l.data_hora_emprestimo
        FROM loans l
        JOIN notebooks n ON n.id = l.notebook_id
        WHERE l.status = 'emprestado'   
        ORDER BY l.data_hora_emprestimo DESC
        """
    )
    rows_ativos = cur.fetchall()

    emprestimos_ativos = []
    for r in rows_ativos:
        d = dict(r)
        texto_limite, nivel_limite = calcular_tempo_limite(d["data_hora_emprestimo"])
        d["tempo_limite"] = texto_limite
        d["nivel_limite"] = nivel_limite
        emprestimos_ativos.append(d)

    total_emprestados = len(emprestimos_ativos)
    disponiveis = total_notebooks - total_emprestados

    # Agendamentos
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        SELECT
            s.id,
            n.serial,
            n.modelo,
            s.colaborador_nome,
            s.colaborador_setor,
            s.data_inicio,
            s.data_fim,
            s.inclui_som
        FROM schedules s
        JOIN notebooks n ON n.id = s.notebook_id
        WHERE s.status = 'agendado'
          AND s.data_fim >= ?
        ORDER BY s.data_inicio
        """,
        (agora,),
    )
    rows_sched = cur.fetchall()

    proximos_agendamentos = []
    hoje = datetime.now().date()
    amanha = hoje + timedelta(days=1)

    for r in rows_sched:
        d = dict(r)
        dt_i = datetime.strptime(d["data_inicio"], "%Y-%m-%d %H:%M:%S")
        dt_f = datetime.strptime(d["data_fim"], "%Y-%m-%d %H:%M:%S")
        d["data_inicio_br"] = dt_i.strftime("%d/%m/%Y %H:%M")
        d["data_fim_br"] = dt_f.strftime("%d/%m/%Y %H:%M")

        if dt_i.date() == hoje:
            d["dia_label"] = "HOJE"
            d["classe_dia"] = "hoje"
        elif dt_i.date() == amanha:
            d["dia_label"] = "AMANHÃ"
            d["classe_dia"] = "amanha"
        else:
            d["dia_label"] = dt_i.strftime("%d/%m")
            d["classe_dia"] = "futuro"

        proximos_agendamentos.append(d)

    # Chamados em aberto
    cur.execute(
        """
        SELECT
            id,
            tipo_equipamento,
            numero_serie,
            empresa,
            data_abertura,
            status
        FROM tickets
        WHERE status = 'aberto'
        ORDER BY data_abertura
        """
    )
    rows_tickets = cur.fetchall()

    chamados_abertos = []
    for r in rows_tickets:
        d = dict(r)
        dt_ab = datetime.strptime(d["data_abertura"], "%Y-%m-%d %H:%M:%S")
        d["data_abertura_br"] = dt_ab.strftime("%d/%m/%Y %H:%M")
        dias = (hoje - dt_ab.date()).days
        d["dias_aberto"] = dias
        chamados_abertos.append(d)

    # Rotas para TV (últimas 10)
    cur.execute(
        """
        SELECT
            id,
            data_solicitacao,
            solicitante,
            unidade_origem,
            prioridade,
            destino,
            descricao_volume,
            status
        FROM rotas
        WHERE status = 'pendente'
        ORDER BY data_solicitacao DESC
        LIMIT 10
        """
    )
    rows_rotas_tv = cur.fetchall()

    rotas_tv = []
    for r in rows_rotas_tv:
        d = dict(r)
        dt = datetime.strptime(d["data_solicitacao"], "%Y-%m-%d %H:%M:%S")
        d["data_solicitacao_br"] = dt.strftime("%d/%m/%Y %H:%M")

        label, css = mapear_prioridade(d.get("prioridade", ""))
        d["prioridade_label"] = label
        d["prioridade_css"] = css

        # <-- ESSA LINHA TEM QUE ESTAR DENTRO DO FOR
        rotas_tv.append(d)

    conn.close()

    return render_template(
        "tv.html",
        body_class="tv-mode",
        total_notebooks=total_notebooks,
        total_emprestados=total_emprestados,
        disponiveis=disponiveis,
        emprestimos_ativos=emprestimos_ativos,
        proximos_agendamentos=proximos_agendamentos,
        chamados_abertos=chamados_abertos,
        rotas_tv=rotas_tv,
    )



@app.route("/relatorio/emprestimo")
@tecnico_required
def relatorio_emprestimo():
    data_inicio = request.args.get("data_inicio")
    data_fim = request.args.get("data_fim")
    setor = request.args.get("setor") or ""
    status = request.args.get("status") or "todos"
    export = request.args.get("export")  # emprestimos / agendamentos / chamados / rotas / all

    hoje = datetime.now().date()

    # Se não vier data, usa últimos 7 dias
    if data_inicio:
        dt_ini = datetime.strptime(data_inicio, "%Y-%m-%d").date()
    else:
        dt_ini = hoje - timedelta(days=7)

    if data_fim:
        dt_fim = datetime.strptime(data_fim, "%Y-%m-%d").date()
    else:
        dt_fim = hoje

    # strings para o form
    data_inicio_str = dt_ini.strftime("%Y-%m-%d")
    data_fim_str = dt_fim.strftime("%Y-%m-%d")

    # limites para SQL (dia inteiro)
    dt_ini_sql = f"{data_inicio_str} 00:00:00"
    dt_fim_sql = f"{data_fim_str} 23:59:59"

    conn = get_connection()
    cur = conn.cursor()

    # =========================
    # 1) EMPRÉSTIMOS
    # =========================
    where = ["l.data_hora_emprestimo BETWEEN ? AND ?"]
    params = [dt_ini_sql, dt_fim_sql]

    if setor:
        where.append("l.colaborador_setor = ?")
        params.append(setor)

    if status and status != "todos":
        where.append("l.status = ?")
        params.append(status)

    sql_loans = f"""
        SELECT
            l.id,
            n.serial,
            n.modelo,
            l.colaborador_nome,
            l.colaborador_setor,
            l.data_hora_emprestimo,
            l.data_hora_devolucao,
            l.status
        FROM loans l
        JOIN notebooks n ON n.id = l.notebook_id
        WHERE {" AND ".join(where)}
        ORDER BY l.data_hora_emprestimo DESC
    """
    cur.execute(sql_loans, params)
    rows_loans = cur.fetchall()

    # =========================
    # 2) AGENDAMENTOS
    # =========================
    cur.execute(
        """
        SELECT
            s.id,
            n.serial,
            n.modelo,
            s.colaborador_nome,
            s.colaborador_setor,
            s.data_inicio,
            s.data_fim,
            s.inclui_som,
            s.status
        FROM schedules s
        JOIN notebooks n ON n.id = s.notebook_id
        WHERE s.data_inicio BETWEEN ? AND ?
        ORDER BY s.data_inicio DESC
        """,
        (dt_ini_sql, dt_fim_sql),
    )
    rows_ag = cur.fetchall()

    # =========================
    # 3) CHAMADOS
    # =========================
    cur.execute(
        """
        SELECT
            id,
            tipo_equipamento,
            numero_serie,
            empresa,
            marca,
            modelo,
            defeito,
            data_abertura,
            status
        FROM tickets
        WHERE data_abertura BETWEEN ? AND ?
        ORDER BY data_abertura DESC
        """,
        (dt_ini_sql, dt_fim_sql),
    )
    rows_ch = cur.fetchall()

    # =========================
    # 4) ROTAS
    # =========================
    cur.execute(
        """
        SELECT
            id,
            data_solicitacao,
            solicitante,
            unidade_origem,
            prioridade,
            destino,
            descricao_volume,
            status
        FROM rotas
        WHERE data_solicitacao BETWEEN ? AND ?
        ORDER BY data_solicitacao DESC
        """,
        (dt_ini_sql, dt_fim_sql),
    )
    rows_rotas = cur.fetchall()

    conn.close()

    # ---------- Monta listas para tela + CSV ----------

    # Empréstimos
    emprestimos = []
    for r in rows_loans:
        d = dict(r)
        dt_e = datetime.strptime(d["data_hora_emprestimo"], "%Y-%m-%d %H:%M:%S")
        d["emprestimo_br"] = dt_e.strftime("%d/%m/%Y %H:%M")
        if d["data_hora_devolucao"]:
            dt_d = datetime.strptime(d["data_hora_devolucao"], "%Y-%m-%d %H:%M:%S")
            d["devolucao_br"] = dt_d.strftime("%d/%m/%Y %H:%M")
        else:
            d["devolucao_br"] = "-"
        emprestimos.append(d)

    # Agendamentos
    agendamentos = []
    for r in rows_ag:
        d = dict(r)
        di = datetime.strptime(d["data_inicio"], "%Y-%m-%d %H:%M:%S")
        df = datetime.strptime(d["data_fim"], "%Y-%m-%d %H:%M:%S")
        d["data_inicio_br"] = di.strftime("%d/%m/%Y %H:%M")
        d["data_fim_br"] = df.strftime("%d/%m/%Y %H:%M")
        agendamentos.append(d)

    # Chamados
    chamados = []
    for r in rows_ch:
        d = dict(r)
        dt_ab = datetime.strptime(d["data_abertura"], "%Y-%m-%d %H:%M:%S")
        d["data_abertura_br"] = dt_ab.strftime("%d/%m/%Y %H:%M")
        d["dias_aberto"] = (hoje - dt_ab.date()).days
        chamados.append(d)

    # Rotas
    rotas = []
    for r in rows_rotas:
        d = dict(r)
        dt_rs = datetime.strptime(d["data_solicitacao"], "%Y-%m-%d %H:%M:%S")
        d["data_solicitacao_br"] = dt_rs.strftime("%d/%m/%Y %H:%M")
        label, css = mapear_prioridade(d.get("prioridade", ""))
        d["prioridade_label"] = label
        d["prioridade_css"] = css
        rotas.append(d)

    # ===== Helpers para gerar CSV =====
    def csv_emprestimos():
        si = io.StringIO()
        w = csv.writer(si, delimiter=";")
        w.writerow(
            [
                "Nº Série",
                "Modelo",
                "Colaborador",
                "Setor",
                "Empréstimo",
                "Devolução",
                "Status",
            ]
        )
        for e in emprestimos:
            w.writerow(
                [
                    e["serial"],
                    e["modelo"],
                    e["colaborador_nome"],
                    e["colaborador_setor"],
                    e["emprestimo_br"],
                    e["devolucao_br"],
                    e["status"],
                ]
            )
        return si.getvalue()

    def csv_agendamentos():
        si = io.StringIO()
        w = csv.writer(si, delimiter=";")
        w.writerow(
            [
                "Nº Série",
                "Modelo",
                "Colaborador",
                "Setor",
                "Início",
                "Fim",
                "Som",
                "Status",
            ]
        )
        for a in agendamentos:
            w.writerow(
                [
                    a["serial"],
                    a["modelo"],
                    a["colaborador_nome"],
                    a["colaborador_setor"],
                    a["data_inicio_br"],
                    a["data_fim_br"],
                    "Com som" if a["inclui_som"] else "Sem som",
                    a["status"],
                ]
            )
        return si.getvalue()

    def csv_chamados():
        si = io.StringIO()
        w = csv.writer(si, delimiter=";")
        w.writerow(
            [
                "Tipo equipamento",
                "Nº Série",
                "Empresa",
                "Marca",
                "Modelo",
                "Defeito",
                "Data abertura",
                "Status",
                "Dias em aberto",
            ]
        )
        for c in chamados:
            w.writerow(
                [
                    c["tipo_equipamento"],
                    c["numero_serie"],
                    c["empresa"],
                    c["marca"],
                    c["modelo"],
                    c["defeito"],
                    c["data_abertura_br"],
                    c["status"],
                    c["dias_aberto"],
                ]
            )
        return si.getvalue()

    def csv_rotas():
        si = io.StringIO()
        w = csv.writer(si, delimiter=";")
        w.writerow(
            [
                "Data solicitação",
                "Solicitante",
                "Unidade de origem",
                "Prioridade",
                "Destino",
                "Descrição do volume",
                "Status",
            ]
        )
        for r in rotas:
            w.writerow(
                [
                    r["data_solicitacao_br"],
                    r["solicitante"],
                    r["unidade_origem"],
                    r["prioridade_label"],
                    r["destino"],
                    r["descricao_volume"],
                    r["status"],
                ]
            )
        return si.getvalue()

    # =========================
    # EXPORT CSV / ZIP
    # =========================
    if export:
        # Export individual
        if export == "emprestimos":
            csv_data = csv_emprestimos()
            resp = make_response(csv_data)
            resp.headers["Content-Type"] = "text/csv; charset=utf-8"
            resp.headers[
                "Content-Disposition"
            ] = "attachment; filename=relatorio_emprestimos.csv"
            return resp

        if export == "agendamentos":
            csv_data = csv_agendamentos()
            resp = make_response(csv_data)
            resp.headers["Content-Type"] = "text/csv; charset=utf-8"
            resp.headers[
                "Content-Disposition"
            ] = "attachment; filename=relatorio_agendamentos.csv"
            return resp

        if export == "chamados":
            csv_data = csv_chamados()
            resp = make_response(csv_data)
            resp.headers["Content-Type"] = "text/csv; charset=utf-8"
            resp.headers[
                "Content-Disposition"
            ] = "attachment; filename=relatorio_chamados.csv"
            return resp

        if export == "rotas":
            csv_data = csv_rotas()
            resp = make_response(csv_data)
            resp.headers["Content-Type"] = "text/csv; charset=utf-8"
            resp.headers[
                "Content-Disposition"
            ] = "attachment; filename=relatorio_rotas.csv"
            return resp

        # Export TUDO (ZIP com 4 CSVs)
        if export == "all":
            mem_zip = io.BytesIO()
            with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(
                    "relatorio_emprestimos.csv",
                    csv_emprestimos().encode("utf-8-sig"),
                )
                zf.writestr(
                    "relatorio_agendamentos.csv",
                    csv_agendamentos().encode("utf-8-sig"),
                )
                zf.writestr(
                    "relatorio_chamados.csv",
                    csv_chamados().encode("utf-8-sig"),
                )
                zf.writestr(
                    "relatorio_rotas.csv",
                    csv_rotas().encode("utf-8-sig"),
                )

            mem_zip.seek(0)
            resp = make_response(mem_zip.getvalue())
            resp.headers["Content-Type"] = "application/zip"
            resp.headers[
                "Content-Disposition"
            ] = "attachment; filename=relatorios_completos.zip"
            return resp

    # Se não for export, renderiza a página normal
    return render_template(
        "relatorio_emprestimo.html",
        emprestimos=emprestimos,
        agendamentos=agendamentos,
        chamados=chamados,
        rotas=rotas,
        data_inicio=data_inicio_str,
        data_fim=data_fim_str,
        setor=setor,
        status=status,
    )

# ===================== GESTÃO DE USUÁRIOS (CONFIGURAÇÕES) =====================

@app.route("/usuarios", methods=["GET", "POST"])
@admin_required
def gerenciar_usuarios():
    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":
        username = request.form.get("username").strip()
        nome_completo = request.form.get("nome_completo").strip()
        role = request.form.get("role")
        
        # A SENHA AGORA É FIXA (REMOVIDO DO FORM)
        senha_padrao = "Rios@ude1234"

        if not username or not nome_completo or not role:
            flash("Preencha todos os campos.", "error")
        else:
            try:
                hashed_pw = generate_password_hash(senha_padrao)
                cur.execute(
                    "INSERT INTO users (username, password_hash, nome_completo, role, must_change_password) VALUES (?, ?, ?, ?, 1)",
                    (username, hashed_pw, nome_completo, role)
                )
                conn.commit()
                flash(f"Usuário cadastrado com sucesso! A senha inicial é {senha_padrao}", "success")
            except sqlite3.IntegrityError:
                flash("Este nome de usuário (login) já existe no sistema.", "error")

    cur.execute("SELECT id, username, nome_completo, role FROM users ORDER BY role, nome_completo")
    usuarios = cur.fetchall()
    conn.close()

    return render_template("usuarios.html", usuarios=usuarios)

# ROTA PARA RESETAR SENHA (ADICIONE ABAIXO DA DELETAR)
@app.route("/usuarios/resetar/<int:user_id>", methods=["POST"])
@admin_required
def resetar_senha_usuario(user_id):
    conn = get_connection()
    cur = conn.cursor()
    senha_padrao = "Rios@ude1234"
    hashed_pw = generate_password_hash(senha_padrao)
    
    cur.execute(
        "UPDATE users SET password_hash = ?, must_change_password = 1 WHERE id = ?", 
        (hashed_pw, user_id)
    )
    conn.commit()
    conn.close()
    flash("A senha do usuário foi redefinida para 'Rios@ude1234' com sucesso.", "success")
    return redirect(url_for("gerenciar_usuarios"))

@app.route("/usuarios/remover/<int:user_id>", methods=["POST"])
@admin_required
def remover_usuario(user_id):
    if user_id == session.get("user_id"):
        flash("Você não pode excluir a sua própria conta logada.", "error")
        return redirect(url_for("gerenciar_usuarios"))
        
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()
    flash("Usuário removido com sucesso.", "success")
    return redirect(url_for("gerenciar_usuarios"))
 

      
if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=False)

