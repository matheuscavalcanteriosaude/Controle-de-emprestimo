import imaplib
import email
from email.header import decode_header
import sqlite3
from datetime import datetime
import re

# =============== CONFIGURAÇÕES ===============

IMAP_SERVER = "imap.gmail.com"
IMAP_USER = "dgovisuporte.riosaude@prefeitura.rio"
IMAP_PASSWORD = "psvf dyhy drqo tgqg"

# mesmo banco usado pelo app Flask
APP_DB = "notebooks.db"

# Assunto exato do formulário de rotas
ASSUNTO_ROTA = (
    "Agradecemos o preenchimento deste formulário: "
    "Solicitação de Transporte para Rotas - RioSaúde"
)

# ============================================


def decode_mime_words(s: str) -> str:
    """Decodifica cabeçalhos MIME (ex.: Subject)."""
    if not s:
        return ""
    decoded = decode_header(s)
    parts = []
    for text, enc in decoded:
        if isinstance(text, bytes):
            try:
                parts.append(text.decode(enc or "utf-8", errors="ignore"))
            except LookupError:
                parts.append(text.decode("utf-8", errors="ignore"))
        else:
            parts.append(text)
    return "".join(parts)


def get_email_body(msg) -> str:
    """
    Retorna o corpo do e-mail como texto.
    Tenta primeiro text/plain; se não tiver, usa text/html e tira as tags.
    """
    body = ""

    if msg.is_multipart():
        # 1ª tentativa: text/plain
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition", "")).lower()

            # Ignora anexos
            if "attachment" in disp:
                continue

            if ctype == "text/plain":
                try:
                    body_bytes = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    body = body_bytes.decode(charset, errors="ignore")
                    return body
                except Exception:
                    continue

        # 2ª tentativa: text/html
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition", "")).lower()
            if "attachment" in disp:
                continue
            if ctype == "text/html":
                try:
                    body_bytes = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    html = body_bytes.decode(charset, errors="ignore")
                    # remove tags HTML básicas para virar texto
                    text = re.sub(r"<br\s*/?>", "\n", html, flags=re.I)
                    text = re.sub(r"<[^>]+>", " ", text)
                    text = re.sub(r"&nbsp;", " ", text)
                    return text
                except Exception:
                    continue
    else:
        ctype = msg.get_content_type()
        body_bytes = msg.get_payload(decode=True)
        if body_bytes:
            charset = msg.get_content_charset() or "utf-8"
            text = body_bytes.decode(charset, errors="ignore")
            if ctype == "text/html":
                text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
                text = re.sub(r"<[^>]+>", " ", text)
                text = re.sub(r"&nbsp;", " ", text)
            return text

    return body or ""


def extrair_descricao_volume(body: str) -> str:
    """
    Extrai a resposta do campo 'Descrição do Volume *' do texto do e-mail.

    Estrutura típica no e-mail:

    Descrição do Volume *
    Descreva detalhadamente o que está sendo enviado (ex: Dipirona, ampola - 200
    unidades)

    Pulseiras de identificação de paciente
    """

    linhas = [l.strip() for l in body.splitlines()]

    # 1) Acha a linha "Descrição do Volume *"
    idx_label = None
    for i, linha in enumerate(linhas):
        if linha.startswith("Descrição do Volume"):
            idx_label = i
            break

    if idx_label is None:
        return ""

    j = idx_label + 1

    # 2) Pula linhas vazias logo após o título
    while j < len(linhas) and not linhas[j]:
        j += 1

    # 3) Se a próxima linha for a de instrução ("Descreva detalhadamente..."),
    #    pula TODA a instrução (que pode ocupar várias linhas) até uma linha vazia
    if j < len(linhas) and linhas[j].startswith("Descreva detalhadamente"):
        j += 1
        # ainda parte da instrução (ex.: "unidades)")
        while j < len(linhas) and linhas[j]:
            j += 1
        # agora pula linhas vazias depois da instrução
        while j < len(linhas) and not linhas[j]:
            j += 1

    # 4) A partir daqui, são as linhas da resposta do usuário.
    descr_linhas = []
    while j < len(linhas):
        t = linhas[j].strip()

        # fim da descrição: linha vazia ou início de outra seção
        if not t:
            break
        if t.startswith("Em caso de rota tipo"):
            break

        descr_linhas.append(t)
        j += 1

    return " ".join(descr_linhas).strip()



def extrair_campos_rota(body: str) -> dict:
    """
    Lê o texto do e-mail do Google Forms e extrai:
    - solicitante
    - unidade_origem
    - prioridade
    - destino
    - descricao_volume
    """
    linhas = [l.strip() for l in body.splitlines()]
    campos = {
        "solicitante": "",
        "unidade_origem": "",
        "prioridade": "",
        "destino": "",
        "descricao_volume": "",
    }

    def proxima_linha_nao_vazia(idx: int) -> str:
        j = idx + 1
        while j < len(linhas) and linhas[j] == "":
            j += 1
        return linhas[j] if j < len(linhas) else ""

    for i, linha in enumerate(linhas):
        if linha.startswith("Nome do Profissional solicitante"):
            campos["solicitante"] = proxima_linha_nao_vazia(i)

        elif linha.startswith("Unidade do solicitante da rota"):
            unidade = proxima_linha_nao_vazia(i)
            # regra especial: Sede Administrativa da RioSaúde = DGOVI
            if "Sede Administrativa da RioSaúde" in unidade:
                campos["unidade_origem"] = "DGOVI"
            else:
                campos["unidade_origem"] = unidade

        elif linha.startswith("Prioridade da Rota"):
            campos["prioridade"] = proxima_linha_nao_vazia(i)

        elif linha.startswith("Destino da Rota"):
            campos["destino"] = proxima_linha_nao_vazia(i)

        elif linha.startswith("Descrição do Volume"):
            # usa a função GLOBAL abaixo
            campos["descricao_volume"] = extrair_descricao_volume(body)

    return campos


def inserir_rota_db(campos: dict):
    """
    Insere uma linha na tabela 'rotas' do notebooks.db.
    """
    conn = sqlite3.connect(APP_DB)
    cur = conn.cursor()

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
            campos["solicitante"],
            campos["unidade_origem"],
            campos["prioridade"],
            campos["destino"],
            campos["descricao_volume"],
        ),
    )

    conn.commit()
    conn.close()


def processar_rotas_email():
    # Conecta no IMAP
    imap = imaplib.IMAP4_SSL(IMAP_SERVER)
    imap.login(IMAP_USER, IMAP_PASSWORD)

    # Caixa de entrada
    imap.select("INBOX")

    # 🔹 Busca APENAS e-mails não lidos
    status, data = imap.search(None, "UNSEEN")

    if status != "OK":
        print("Falha ao buscar emails.")
        imap.logout()
        return

    ids = data[0].split()
    print(f"{len(ids)} e-mails não lidos encontrados.")

    for num in ids:
        # Busca o e-mail completo
        status, msg_data = imap.fetch(num, "(RFC822)")
        if status != "OK":
            print(f"Falha ao buscar a mensagem {num}.")
            continue

        msg = email.message_from_bytes(msg_data[0][1])

        # Assunto decodificado
        subject = decode_mime_words(msg.get("Subject", ""))
        print(f"Assunto: {subject}")

        # Garante que é o assunto certo do formulário de rotas
        if ASSUNTO_ROTA not in subject:
            print("  → Assunto não corresponde ao formulário de rotas. Ignorando.")
            # NÃO marca como lido
            continue

        # Corpo
        body = get_email_body(msg)

        # Log opcional do corpo
        print("===== INÍCIO BODY =====")
        print(body)
        print("===== FIM BODY =====")

        campos = extrair_campos_rota(body)

        obrigatorios = [
            "solicitante",
            "unidade_origem",
            "prioridade",
            "destino",
            "descricao_volume",
        ]
        faltando = [k for k in obrigatorios if not campos.get(k)]

        if faltando:
            print(f"  → Email ignorado, campos obrigatórios faltando: {faltando}")
            # NÃO marca como lido, pra você poder ver depois
            continue

        # Insere no banco
        inserir_rota_db(campos)
        print(
            f"  → Rota inserida no banco para {campos['solicitante']} "
            f"({campos['unidade_origem']})"
        )

        # Agora sim, marca como lido, já que foi processado com sucesso
        imap.store(num, "+FLAGS", "\\Seen")

    imap.close()
    imap.logout()


if __name__ == "__main__":
    processar_rotas_email()
