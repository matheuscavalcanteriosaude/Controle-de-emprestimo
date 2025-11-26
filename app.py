import sqlite3
import csv
import subprocess
import shutil
import os
import re
from io import StringIO
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    Response,
    send_file,
    abort,

)

APP_DB = "notebooks.db"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TERMOS_DIR = os.path.join(BASE_DIR, "termos")
TERMOS_GERADOS_DIR = os.path.join(BASE_DIR, "termos_gerados")
os.makedirs(TERMOS_GERADOS_DIR, exist_ok=True)

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

    # Tabela de empréstimos
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
    return redirect(url_for("dashboard"))

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




# ===================== CADASTRO / LISTA DE NOTEBOOKS =====================

@app.route("/notebooks/novo", methods=["GET", "POST"])
def novo_notebook():
    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":
        serial = request.form.get("serial", "").strip()
        modelo = request.form.get("modelo", "").strip()

        if not serial:
            flash("O número de série é obrigatório.", "error")
            conn.close()
            return redirect(url_for("novo_notebook"))

        try:
            cur.execute(
                "INSERT INTO notebooks (serial, modelo, status) VALUES (?, ?, 'ativo')",
                (serial, modelo),
            )
            conn.commit()
            flash("Notebook cadastrado com sucesso!", "success")
        except sqlite3.IntegrityError:
            flash("Já existe um notebook com esse número de série.", "error")

    # Lista notebooks (ativos e inativos)
    cur.execute(
        """
        SELECT id, serial, modelo, status
        FROM notebooks
        ORDER BY status DESC, serial
        """
    )
    notebooks = cur.fetchall()
    conn.close()

    return render_template("notebook_form.html", notebooks=notebooks)


@app.route("/notebooks/remover/<int:notebook_id>", methods=["POST"])
def remover_notebook(notebook_id):
    conn = get_connection()
    cur = conn.cursor()

    # Verifica se há empréstimo ativo
    cur.execute(
        "SELECT COUNT(*) AS total FROM loans WHERE notebook_id = ? AND status = 'emprestado'",
        (notebook_id,),
    )
    em_uso = cur.fetchone()["total"]

    if em_uso > 0:
        flash("Não é possível remover: notebook está com empréstimo ativo.", "error")
        conn.close()
        return redirect(url_for("novo_notebook"))

    cur.execute("UPDATE notebooks SET status = 'inativo' WHERE id = ?", (notebook_id,))
    conn.commit()
    conn.close()

    flash("Notebook removido (marcado como inativo).", "success")
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

    # Total de notebooks ativos (estoque)
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

    # Histórico recente (últimos 10)
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

    # Próximos agendamentos (para o dashboard)
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
    conn.close()

    proximos_agendamentos = []
    hoje = datetime.now().date()
    for r in rows_sched:
        d = dict(r)
        dt_i = datetime.strptime(d["data_inicio"], "%Y-%m-%d %H:%M:%S")
        dt_f = datetime.strptime(d["data_fim"], "%Y-%m-%d %H:%M:%S")
        d["data_inicio_br"] = dt_i.strftime("%d/%m/%Y %H:%M")
        d["data_fim_br"] = dt_f.strftime("%d/%m/%Y %H:%M")
        d["eh_hoje"] = (dt_i.date() == hoje)
        proximos_agendamentos.append(d)

    return render_template(
        "dashboard.html",
        total_notebooks=total_notebooks,
        total_emprestados=total_emprestados,
        disponiveis=disponiveis,
        emprestimos_ativos=emprestimos_ativos,
        historico=historico,
        proximos_agendamentos=proximos_agendamentos,
    )

@app.route("/relatorio/emprestimo")
def relatorio_emprestimo():
    status = request.args.get("status", "todos")
    setor = request.args.get("setor", "").strip()
    dt_ini = request.args.get("data_inicio", "").strip()
    dt_fim = request.args.get("data_fim", "").strip()
    export = request.args.get("export")

    conn = get_connection()
    cur = conn.cursor()

    conds = []
    params = []

    if status in ("emprestado", "devolvido"):
        conds.append("l.status = ?")
        params.append(status)

    if setor:
        conds.append("l.colaborador_setor LIKE ?")
        params.append("%" + setor + "%")

    if dt_ini:
        # dt_ini vem como YYYY-MM-DD do input type="date"
        conds.append("date(l.data_hora_emprestimo) >= date(?)")
        params.append(dt_ini)

    if dt_fim:
        conds.append("date(l.data_hora_emprestimo) <= date(?)")
        params.append(dt_fim)

    where = ""
    if conds:
        where = "WHERE " + " AND ".join(conds)

    query = f"""
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
        {where}
        ORDER BY l.data_hora_emprestimo DESC
    """

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    # Se for exportar CSV
    if export == "csv":
        si = StringIO()
        writer = csv.writer(si, delimiter=';')

        # Cabeçalho
        writer.writerow([
            "ID",
            "Nº Série",
            "Modelo",
            "Colaborador",
            "Setor",
            "Data/Hora Empréstimo",
            "Data/Hora Devolução",
            "Status",
        ])

        # Linhas
        for r in rows:
            writer.writerow([
                r["id"],
                r["serial"],
                r["modelo"] or "",
                r["colaborador_nome"],
                r["colaborador_setor"] or "",
                r["data_hora_emprestimo"],
                r["data_hora_devolucao"] or "",
                r["status"],
            ])

        output = si.getvalue()
        resp = Response(output, mimetype="text/csv; charset=utf-8")
        resp.headers["Content-Disposition"] = "attachment; filename=relatorio_emprestimo.csv"
        return resp

    # Caso normal: exibir na tela
    return render_template(
        "relatorio_emprestimo.html",
        rows=rows,
        status=status,
        setor=setor,
        data_inicio=dt_ini,
        data_fim=dt_fim,
    )

# ===================== MODO TV =====================

@app.route("/tv")
def tv_dashboard():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) AS total FROM notebooks WHERE status = 'ativo'")
    total_notebooks = cur.fetchone()["total"]

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
    conn.close()

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

    return render_template(
        "tv.html",
        body_class="tv-mode",
        total_notebooks=total_notebooks,
        total_emprestados=total_emprestados,
        disponiveis=disponiveis,
        emprestimos_ativos=emprestimos_ativos,
        proximos_agendamentos=proximos_agendamentos,
    )


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=False)

