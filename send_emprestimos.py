# send_prazos_emprestimos.py

import os
import sqlite3
import datetime
import requests
from pathlib import Path

# =========================
# CONFIGURAÇÃO TELEGRAM
# =========================
# Use os MESMOS valores que você já usa no send_agendamentos_hoje.py
TELEGRAM_TOKEN = "8086550452:AAFTGEn8hQ8wWkStiF_n0BJ4GozSfiCcQEs"
CHAT_ID = "-5184416502"  # pode ser usuário ou grupo

# =========================
# CONFIGURAÇÃO DO BANCO
# =========================
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "notebooks.db"

# Limite de empréstimo em horas (18h)
LIMITE_HORAS = 18

# Janela para "quase estourando" (ex.: 1h)
JANELA_ALERTA_HORAS = 1


# =========================
# FUNÇÕES AUXILIARES
# =========================
def parse_datetime(dt_str: str) -> datetime.datetime:
    """
    Converte a string do banco para datetime.
    Espera algo como '2025-12-09 14:54:34'.
    """
    if not dt_str:
        return None
    # fromisoformat aceita 'YYYY-MM-DD HH:MM:SS'
    return datetime.datetime.fromisoformat(dt_str)


def format_timedelta(td: datetime.timedelta) -> str:
    """
    Recebe um timedelta e devolve texto tipo '1h 30min'.
    Usa o valor absoluto (sem sinal).
    """
    total_min = int(abs(td.total_seconds()) // 60)
    horas = total_min // 60
    minutos = total_min % 60

    if horas > 0 and minutos > 0:
        return f"{horas}h {minutos}min"
    elif horas > 0:
        return f"{horas}h"
    else:
        return f"{minutos}min"


# =========================
# BUSCAR EMPRÉSTIMOS
# =========================
def get_emprestimos_criticos():
    """
    Busca empréstimos ativos (status = 'emprestado') e retorna
    apenas aqueles cujo limite de 18h está:
      - a <= 2h de estourar, ou
      - já estourou.
    """
    if not DB_PATH.exists():
        print(f"[DEBUG] Banco não encontrado em {DB_PATH}")
        return []

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute(
        """
        SELECT
            l.id,
            l.notebook_id,
            l.colaborador_nome,
            l.colaborador_setor,
            l.data_hora_emprestimo,
            l.data_hora_devolucao,
            l.status,
            n.serial AS notebook_serial,
            n.modelo AS notebook_modelo
        FROM loans l
        JOIN notebooks n ON n.id = l.notebook_id
        WHERE l.status = 'emprestado'
        """
    )
    rows = cur.fetchall()
    conn.close()

    agora = datetime.datetime.now()
    resultados = []

    for row in rows:
        dt_emp = parse_datetime(row["data_hora_emprestimo"])
        if not dt_emp:
            continue

        limite = dt_emp + datetime.timedelta(hours=LIMITE_HORAS)
        delta = limite - agora  # positivo = ainda falta; negativo = atrasado

        # Já atrasado?
        if delta.total_seconds() < 0:
            situacao = "ATRASADO"
            tempo = agora - limite
            tempo_txt = format_timedelta(tempo)
        else:
            # Falta mais que a janela? então ainda não é crítico
            if delta > datetime.timedelta(hours=JANELA_ALERTA_HORAS):
                continue
            situacao = "PRÓXIMO DO LIMITE"
            tempo_txt = format_timedelta(delta)

        resultados.append(
            {
                "id": row["id"],
                "serial": row["notebook_serial"],
                "modelo": row["notebook_modelo"],
                "colaborador": row["colaborador_nome"],
                "setor": row["colaborador_setor"],
                "data_emprestimo": dt_emp,
                "limite": limite,
                "situacao": situacao,
                "tempo_txt": tempo_txt,
            }
        )

    return resultados


# =========================
# MONTAR MENSAGEM
# =========================
def montar_mensagem(emprestimos):
    if not emprestimos:
        return None  # nada crítico, não manda

    agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    partes = [
        f"⚠️ Aviso de prazo de empréstimo\n{agora}\n",
        f"Limite padrão: {LIMITE_HORAS}h após a retirada.\n",
    ]

    for i, emp in enumerate(emprestimos, start=1):
        serial = emp["serial"]
        modelo = emp["modelo"]
        colab = emp["colaborador"]
        setor = emp["setor"]
        dt_emp = emp["data_emprestimo"].strftime("%d/%m/%Y %H:%M")
        limite = emp["limite"].strftime("%d/%m/%Y %H:%M")
        situacao = emp["situacao"]
        tempo_txt = emp["tempo_txt"]

        if situacao == "ATRASADO":
            linha_sit = f"{situacao} há {tempo_txt} ⛔"
        else:
            linha_sit = f"{situacao}: falta {tempo_txt} ⏰"

        partes.append(
            (
                f"\n{i}. {serial} – {modelo}\n"
                f"   {colab} ({setor})\n"
                f"   Empréstimo: {dt_emp}\n"
                f"   Limite: {limite}\n"
                f"   {linha_sit}"
            )
        )

    return "\n".join(partes)


# =========================
# ENVIAR TELEGRAM
# =========================
def enviar_telegram(texto):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {"chat_id": CHAT_ID, "text": texto}

    print(f"[DEBUG] Enviando POST para Telegram: {url}")
    print(f"[DEBUG] Payload: {payload}")

    resp = requests.post(url, json=payload)
    print(f"[DEBUG] Status code Telegram: {resp.status_code}")
    try:
        print(f"[DEBUG] Resposta Telegram: {resp.text}")
    except Exception:
        pass

    resp.raise_for_status()


# =========================
# MAIN
# =========================
def main():
    print("=== Script send_prazos_emprestimos INICIADO ===")
    print(f"[DEBUG] Usando DB_PATH: {DB_PATH}")

    emprestimos = get_emprestimos_criticos()
    print(f"[DEBUG] Qtde de empréstimos críticos encontrados: {len(emprestimos)}")

    mensagem = montar_mensagem(emprestimos)
    print(f"[DEBUG] Mensagem montada: {repr(mensagem)}")

    if mensagem:
        enviar_telegram(mensagem)
    else:
        print("[DEBUG] Nenhuma mensagem enviada (sem empréstimos críticos).")

    print("=== Script send_prazos_emprestimos FINALIZADO ===")


if __name__ == "__main__":
    main()
