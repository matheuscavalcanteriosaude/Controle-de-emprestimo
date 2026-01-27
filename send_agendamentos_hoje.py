import os
import sqlite3
import datetime
import requests

# =========================
# CONFIGURAÇÕES
# =========================
TELEGRAM_TOKEN = "8086550452:AAFTGEn8hQ8wWkStiF_n0BJ4GozSfiCcQEs"
CHAT_ID = "-5184416502"

# Caminho do seu banco (ajuste para o seu caso real)
DB_PATH = r"C:\Users\srsadmin\Desktop\notebooks - online 28-11\notebooks.db"

def get_agendamentos_hoje():
    hoje = datetime.date.today().isoformat()  # 'YYYY-MM-DD'
    print(f"[DEBUG] Data de hoje (filtro): {hoje}")
    print(f"[DEBUG] Usando DB_PATH: {DB_PATH}")
    print(f"[DEBUG] Arquivo de banco existe? {os.path.exists(DB_PATH)}")

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute(
        """
        SELECT
            s.id,
            n.modelo                           AS modelo,
            s.colaborador_nome                 AS colaborador,
            s.colaborador_setor                AS setor,
            substr(s.data_inicio,1,10)         AS data_inicio,
            substr(s.data_inicio,12,5)         AS hora_inicio,
            substr(s.data_fim,12,5)            AS hora_fim,
            s.inclui_som                       AS som,
            s.status                           AS status
        FROM schedules s
        JOIN notebooks n ON n.id = s.notebook_id
        WHERE substr(s.data_inicio,1,10) = ?
          AND s.status != 'cancelado'   -- considera tudo que não foi cancelado
        ORDER BY s.data_inicio
        """,
        (hoje,),
    )

    rows = cur.fetchall()
    conn.close()

    print(f"[DEBUG] Qtde de agendamentos encontrados para hoje: {len(rows)}")
    for r in rows:
        print("[DEBUG] Row:", dict(r))

    return rows

# =========================
# MONTAR MENSAGEM
# =========================
def montar_mensagem(agendamentos):
    if not agendamentos:
        return None  # "caso tenha" → se não tiver nada, não manda

    hoje = datetime.date.today().strftime("%d/%m/%Y")
    partes = [f"📅 Agendamentos de hoje ({hoje}):\n"]

    for i, ag in enumerate(agendamentos, start=1):
        # Como usamos row_factory=sqlite3.Row, dá pra acessar por nome
        modelo = ag["modelo"]
        colab = ag["colaborador"]
        setor = ag["setor"]
        data_ini = ag["data_inicio"]
        hora_ini = ag["hora_inicio"]
        hora_fim = ag["hora_fim"]
        som = ag["som"]

        linha = (
            f"{i}. {hora_ini}–{hora_fim}  |  {modelo}  |  {colab} ({setor})"
        )
        if som:
            linha += "  🔊"
        partes.append(linha)

    return "\n".join(partes)


# =========================
# ENVIAR PARA TELEGRAM
# =========================
def enviar_telegram(texto):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": CHAT_ID,
        "text": texto,
    }
    print(f"[DEBUG] Enviando POST para Telegram: {url}")
    print(f"[DEBUG] Payload: {payload}")

    resp = requests.post(url, json=payload, timeout=15)
    print(f"[DEBUG] Status code Telegram: {resp.status_code}")
    print(f"[DEBUG] Resposta Telegram: {resp.text}")

    resp.raise_for_status()


def main():
    print("=== Script send_agendamentos_hoje INICIADO ===")
    ags = get_agendamentos_hoje()
    print(f"[DEBUG] Lista ags (len={len(ags)}): {ags}")

    mensagem = montar_mensagem(ags)
    print(f"[DEBUG] Mensagem montada: {repr(mensagem)}")

    if mensagem:
        print("[DEBUG] Há mensagem, enviando para o Telegram...")
        enviar_telegram(mensagem)
        print("[DEBUG] Mensagem enviada (se não deu erro acima).")
    else:
        print("[DEBUG] Nenhuma mensagem montada (provavelmente sem agendamento para hoje).")

    print("=== Script send_agendamentos_hoje FINALIZADO ===")


if __name__ == "__main__":
    main()
