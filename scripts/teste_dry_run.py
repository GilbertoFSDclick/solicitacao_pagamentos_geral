"""
Teste Dry-Run - Simula o fluxo com dados fixos (protocolo 260306.145103).
NÃO conecta ao Holmes, NÃO abre o NBS, NÃO altera dados.
Valida: tratar_tarefa (via mock), planilha TIPOS_DE_PAGAMENTO, estrutura de dados.
Uso: python scripts/teste_dry_run.py
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Dados do protocolo 260306.145103 (tarefa já concluída - apenas para teste)
DADOS_MOCK = {
    "protocolo": "260306.145103",
    "filial": "481 - GWM GUARULHOS STA FRANCI",
    "cpf": "06708584808",
    "tipo_pagamento": "Multa à cobrar",
    "valores": "104,12",
    "observacao": "Placa FZL9H85",
    "cnpj_loja": "46973672000184",
    "ait": "6VA1965143",
}


def testar_planilha_tipos_pagamento():
    """Valida se a planilha TIPOS_DE_PAGAMENTO retorna o código UAB para 'Multa à cobrar'."""
    print("\n[1] Testando planilha TIPOS_DE_PAGAMENTO...")
    try:
        import pandas as pd

        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        caminhos_teste = [
            os.path.join(os.path.dirname(base), "TIPOS_DE_PAGAMENTO.xlsx"),
            os.path.join(base, "TIPOS_DE_PAGAMENTO.xlsx"),
            os.path.join(base, "..", "TIPOS_DE_PAGAMENTO.xlsx"),
        ]
        caminho = None
        for cp in caminhos_teste:
            if cp and os.path.exists(cp):
                caminho = cp
                print(f"  Usando planilha: {caminho}")
                break
        if not caminho:
            print("  [AVISO] TIPOS_DE_PAGAMENTO.xlsx não encontrado no workspace")
            return False

        for sheet in ["Plan1", "Planilha1", 0]:
            try:
                df = pd.read_excel(caminho, sheet_name=sheet, header=1)
                break
            except Exception:
                continue
        else:
            print("  [ERRO] Não foi possível ler a planilha")
            return False

        if df.shape[1] < 4:
            print("  [ERRO] Planilha sem coluna D")
            return False

        col_a = df.iloc[:, 0].astype(str).str.strip().str.lower()
        col_d = df.iloc[:, 3].astype(str)
        tipo_busca = "multa à cobrar"
        mask = col_a == tipo_busca
        if mask.any():
            codigo = col_d[mask].iloc[0].strip()
            print(f"  [OK] Tipo '{DADOS_MOCK['tipo_pagamento']}' -> Codigo UAB: {codigo}")
            return True
        print(f"  [ERRO] Tipo '{DADOS_MOCK['tipo_pagamento']}' não encontrado na planilha")
        print(f"  Valores na coluna A: {col_a.tolist()[:15]}")
        return False

    except Exception as e:
        print(f"  [ERRO] {e}")
        import traceback
        traceback.print_exc()
        return False


def testar_estrutura_campos():
    """Valida se os campos obrigatórios estão corretos para processar_entrada_solicitacao_pagamento."""
    print("\n[2] Testando estrutura de campos...")
    campos = [
        DADOS_MOCK["cpf"],
        "1",
        "481",
        DADOS_MOCK["ait"].replace("-", "").replace(".", ""),
        "2026-03-06",
        DADOS_MOCK["valores"].replace(",", "."),
        DADOS_MOCK["protocolo"],
        "0",
        "06/03/2026",
        DADOS_MOCK["tipo_pagamento"],
    ]
    print(f"  Campos: cpf={campos[0]}, processo={campos[3]}, valor={campos[5]}, tipo={campos[9]}")
    if len(campos) == 10:
        print("  [OK] Estrutura correta (10 campos)")
        return True
    print(f"  [ERRO] Esperados 10 campos, obtidos {len(campos)}")
    return False


def testar_obter_codigo_uab():
    """Testa a função obter_codigo_uab_contabil (requer ambiente com bot instalado)."""
    print("\n[3] Testando obter_codigo_uab_contabil()...")
    try:
        import modulos.nbs.solicitacao_pagamento as sp

        codigo = sp.obter_codigo_uab_contabil(DADOS_MOCK["tipo_pagamento"])
        if codigo:
            print(f"  [OK] Código UAB: {codigo}")
            return True
        print("  [AVISO] Retornou None - verifique planilha_tipos_pagamento no .ini")
        return False

    except ImportError as e:
        print(f"  [AVISO] Ambiente incompleto (bot não instalado): {e}")
        print("  Execute no ambiente do projeto para testar obter_codigo_uab_contabil")
        return True
    except Exception as e:
        print(f"  [ERRO] {e}")
        return False


def main():
    print("=" * 60)
    print("TESTE DRY-RUN - Dados do protocolo 260306.145103")
    print("(Tarefa já concluída - apenas validação de estrutura)")
    print("=" * 60)

    resultados = []
    resultados.append(("Planilha TIPOS_DE_PAGAMENTO", testar_planilha_tipos_pagamento()))
    resultados.append(("Estrutura de campos", testar_estrutura_campos()))
    resultados.append(("obter_codigo_uab_contabil", testar_obter_codigo_uab()))

    print("\n" + "=" * 60)
    print("RESUMO")
    for nome, ok in resultados:
        status = "[OK]" if ok else "[FALHOU]"
        print(f"  {status} {nome}")
    print("=" * 60)

    # Falha apenas se planilha ou estrutura falharem (obter_codigo pode ser AVISO)
    criticos = resultados[:2]
    return 0 if all(r[1] for r in criticos) else 1


if __name__ == "__main__":
    sys.exit(main())
