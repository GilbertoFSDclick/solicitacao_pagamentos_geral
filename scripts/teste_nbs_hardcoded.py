"""
Teste NBS com dados HARDCODED - tarefa ja concluida (protocolo 260306.145103).
Nao chama webhook, nao chama Holmes, nao encaminha tarefa.
Abre o NBS e executa o fluxo real de processamento com dados fixos.

Modos:
  --parar-antes  = Para apos login no NBS (NAO cria entrada).
  --ultima-etapa = Executa todo o fluxo, preenche o formulario, para ANTES de Confirmar (NAO cria entrada). Padrao.
  --completo     = Executa o fluxo inteiro (cria entrada no NBS).
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Dados fixos da tarefa ja concluida (protocolo 260306.145103)
DADOS_HARDCODED = {
    "cpf": "06708584808",
    "cod_empresa": "1",
    "cod_filial": "481",
    "processo": "6VA1965143",
    "iniciar_em": "2026-03-06",
    "despesa_pagamento": "104.12",
    "protocolo": "260306.145103",
    "centro_custo": "0",
    "data_vencimento": "06/03/2026",
    "tipo_pagamento": "Multa à cobrar",
}


def main():
    parar_apos_login = "--parar-antes" in sys.argv
    parar_ultima_etapa = "--completo" not in sys.argv and not parar_apos_login

    print("=" * 60)
    print("TESTE NBS - Dados hardcoded (protocolo 260306.145103)")
    print("=" * 60)
    print("Este script NAO chama webhook nem Holmes.")
    if parar_apos_login:
        print("")
        print("MODO: --parar-antes")
        print("  Para apos login no NBS. NAO cria entrada.")
    elif parar_ultima_etapa:
        print("")
        print("MODO: --ultima-etapa (padrao)")
        print("  Executa todo o fluxo, preenche o formulario.")
        print("  Para ANTES de Confirmar. NAO cria entrada.")
    else:
        print("")
        print("MODO: --completo")
        print("  ATENCAO: Executara o fluxo inteiro e criara ENTRADA no NBS!")
    print("=" * 60)
    input("Pressione ENTER para continuar ou Ctrl+C para cancelar...")

    import bot
    import modulos
    import pyautogui

    bot.logger.informar("Iniciando teste NBS com dados hardcoded...")
    x, y = pyautogui.size()
    bot.logger.informar(f"Resolucao: {x} x {y}")

    # [1] Iniciar NBS
    try:
        nbs = modulos.nbs.Sistema(
            *bot.configfile.obter_opcoes("nbs", ["usuario", "senha", "servidor"])
        )
        nbs.inicializar()
    except Exception as e:
        bot.logger.erro(f"Erro ao inicializar NBS: {e}")
        return 1

    if parar_apos_login:
        print("")
        print("=" * 60)
        print("PARADO: NBS aberto e logado. Nenhuma entrada criada.")
        print("Para fluxo ate ultima etapa: rodar_teste_nbs.bat")
        print("Para fluxo completo: rodar_teste_nbs_completo.bat")
        print("=" * 60)
        return 0

    campos_obrigatorios = [
        DADOS_HARDCODED["cpf"],
        DADOS_HARDCODED["cod_empresa"],
        DADOS_HARDCODED["cod_filial"],
        DADOS_HARDCODED["processo"],
        DADOS_HARDCODED["iniciar_em"],
        DADOS_HARDCODED["despesa_pagamento"],
        DADOS_HARDCODED["protocolo"],
        DADOS_HARDCODED["centro_custo"],
        DADOS_HARDCODED["data_vencimento"],
        DADOS_HARDCODED["tipo_pagamento"],
    ]
    dados_processo = {"propriedades_processo": []}

    # [2] Processar entrada (fluxo real no NBS)
    try:
        entrada = modulos.nbs.processar_entrada_solicitacao_pagamento(
            campos_obrigatorios,
            dados_processo,
            parar_antes_confirmar=parar_ultima_etapa,
        )
        if entrada.SUCESSO:
            bot.logger.informar("TESTE CONCLUIDO: Fluxo executado com sucesso.")
        else:
            bot.logger.informar(f"TESTE: {entrada.MENSAGEM}")
    except Exception as e:
        bot.logger.erro(f"Erro no processamento: {e}")
        import traceback
        traceback.print_exc()
        return 1

    print("")
    print("=" * 60)
    print("Teste finalizado. Nenhuma tarefa foi encaminhada no Holmes.")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
