from waitress import serve
from app import app, init_db  # <-- Adicionamos a importação do init_db

if __name__ == '__main__':
    print("Verificando/Atualizando o banco de dados...")
    init_db()  # <-- Força a criação de qualquer tabela nova antes de subir o servidor
    
    print("Servidor de Produção Iniciado na porta 5000...")
    # O host '0.0.0.0' permite que outros PCs da rede acessem
    serve(app, host='0.0.0.0', port=5000, threads=6)