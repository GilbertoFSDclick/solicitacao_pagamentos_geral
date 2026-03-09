"""
Script de validação do Webhook - MODO SEGURO (somente leitura).
Exibe as propriedades retornadas pelo webhook para ajustar a classe Properties.
NÃO remove processos, NÃO encaminha tarefas, NÃO altera dados.
Uso: python scripts/validar_webhook.py
"""
import sys
import os
import json

# Garantir que o projeto está no path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import bot
from modulos import webhook


def main():
    print("=" * 60)
    print("VALIDAÇÃO WEBHOOK - Modo somente leitura")
    print("=" * 60)

    try:
        webhook.checar_conexao_webhook()
        print("[OK] Conexão com webhook estabelecida")
    except AssertionError as e:
        print(f"[ERRO] {e}")
        return 1

    try:
        client = webhook.client_singleton()
        response = client.get("/webhook/holmes", params={"limit": 5, "query": "properties.DMS = 'nbs'"})
        if response.status_code != 200:
            print(f"[ERRO] Status {response.status_code}")
            return 1

        dados = response.json()
        processos = dados.get("processos", [])
        print(f"\nProcessos retornados: {len(processos)}")

        if not processos:
            print("\nNenhum processo encontrado. Verifique se há processos com DMS=nbs no webhook.")
            return 0

        for i, proc in enumerate(processos[:3]):
            print(f"\n--- Processo {i + 1}: {proc.get('id_processo', 'N/A')} ---")
            props = proc.get("dados", {}).get("properties", {})
            print("Propriedades (chaves exatas do webhook):")
            for k, v in sorted(props.items()):
                valor = str(v)[:80] + "..." if len(str(v)) > 80 else v
                print(f"  {repr(k)}: {repr(valor)}")

            print("\nChaves normalizadas (bot.util.normalizar):")
            try:
                normalizadas = {bot.util.normalizar(k): v for k, v in props.items()}
                for k, v in sorted(normalizadas.items()):
                    valor = str(v)[:80] + "..." if len(str(v)) > 80 else v
                    print(f"  {repr(k)}: {repr(valor)}")
            except Exception as e:
                print(f"  (erro ao normalizar: {e})")

        print("\n" + "=" * 60)
        print("A classe Properties em src/webhook.py deve ter campos que correspondam")
        print("às chaves NORMALIZADAS acima. Ajuste os nomes dos campos conforme necessário.")
        print("=" * 60)
        return 0

    except Exception as e:
        print(f"[ERRO] {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
