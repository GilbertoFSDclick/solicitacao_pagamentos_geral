"""
Cria a tabela rpa_controle_solicitacao_pagamento no banco de controle.
Execute a partir da raiz do projeto: python scripts/criar_tabela_banco_controle.py
Requer [banco_controle] url configurada no .ini.
"""
import sys
from pathlib import Path

# Garantir que o projeto está no path
raiz = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(raiz))

import bot

COLUNAS = [
    ("id_processo", "TEXT NOT NULL PRIMARY KEY"),
    ("Empresa_holmes", "TEXT"),
    ("Filial_holmes", "TEXT"),
    ("Cpf_holmes", "TEXT"),
    ("Num_nf_holmes", "TEXT"),
    ("Serie_nf_holmes", "TEXT"),
    ("Emissão_nf_holmes", "TEXT"),
    ("Entrada_nf_holmes", "TEXT"),
    ("Vlr_nf_holmes", "TEXT"),
    ("Protocolo_holmes", "TEXT"),
    ("Centro_custo_uab_holmes", "TEXT"),
    ("Dt_Vencimento_Holmes", "TEXT"),
    ("Codigo_UAB_Contabil", "TEXT"),
    ("data_registro", "TEXT DEFAULT (datetime('now', 'localtime'))"),
]


def main():
    url = bot.configfile.obter_opcao_ou("banco_controle", "url")
    tabela = bot.configfile.obter_opcao_ou("banco_controle", "tabela") or "rpa_controle_solicitacao_pagamento"
    if not url:
        print("Configure [banco_controle] url no .ini antes de executar.")
        sys.exit(1)

    try:
        import sqlalchemy
        from sqlalchemy import text
    except ImportError:
        print("Instale: pip install sqlalchemy")
        sys.exit(1)

    # SQLite: CREATE TABLE IF NOT EXISTS
    if "sqlite" in url.lower():
        cols = ", ".join(f'"{c[0]}" {c[1]}' for c in COLUNAS)
        sql = f"CREATE TABLE IF NOT EXISTS {tabela} ({cols})"
    else:
        # Outros (SQL Server etc.): use o arquivo .sql manualmente
        print("Para bancos além de SQLite, execute o script SQL manualmente.")
        print("Arquivo: scripts/criar_tabela_banco_controle.sql")
        sys.exit(0)

    try:
        engine = sqlalchemy.create_engine(url)
        with engine.connect() as conn:
            conn.execute(text(sql))
            conn.commit()
        print(f"Tabela '{tabela}' criada/verificada com sucesso.")
    except Exception as e:
        print(f"Erro: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
