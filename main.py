"""
Fluxo principal - Lançamento de Solicitações de Pagamento Geral.
Conforme Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral.
"""
import bot
import modulos
import operacoes
import pyautogui
from src.webhook import obter_processos
from src.exceptions import ErroNegocio, ErroTecnico

from typing import Literal

NOME_BOT = bot.configfile.obter_opcao_ou("bot", "nome")
NOME_LOG = ".log"


def notificar_email(tipo: Literal["ERRO", "SUCESSO"], resultados: dict | None = None) -> None:
    """Enviar notificação via e-mail (TE02, TE03, TE04, TE07)."""
    html: str
    mensagem = (
        "Bot executado com sucesso"
        if tipo == "SUCESSO"
        else "A execução da automação resultou em erro"
    )
    destinatarios = bot.configfile.obter_opcoes("email.destinatarios", [str(tipo).lower()])[
        0
    ].split(", ")

    if tipo == "SUCESSO" and resultados:
        table_html = """
            <div style="background-color: #f1f1f1; border-radius: 10px; padding:12px;margin:30px 0;border:1px solid #cdcdcd">
                <table style="width: 100%">
                <thead style="background-color: #62adba; height:45px;color:#fff">
                    <tr>
                    <th>Filial</th>
                    <th>Processo</th>
                    <th>Status</th>
                    <th>Protocolo</th>
                    </tr>
                </thead>
                <tbody>
            """
        for processo, dados in resultados.items():
            table_html += f"<tr><th rowspan='2'>{dados[1]}</th></tr>"
            status = "CONCLUÍDO" if dados[0] else "EXCEÇÃO"
            protocolo = dados[2]
            table_html += f'<tr><td style="text-align:center;font-weight:600">{processo}</td>'
            table_html += f'<td style="text-align:center;font-weight:600">{status}</td>'
            table_html += f'<td style="text-align:center;font-weight:600">{protocolo}</td></tr>'
            table_html += f'<tr><td colspan="4" style="border-bottom: 1px solid #bbbbbb;padding:10px"></td></tr>'
        table_html += """
            </tbody>
            </table>
            </div>
            """
    else:
        table_html = ""

    with open("email.html", encoding="utf-8") as arquivo:
        html = (
            arquivo.read()
            .replace("{0}", f"{NOME_BOT} - {tipo}")
            .replace(
                "{1}",
                f"{mensagem}. Todos os detalhes do processamento estão no log em anexo.",
            )
            .replace("{2}", NOME_BOT)
            .replace("{table_resultados}", table_html)
        )

    assunto = f"AUTOMAÇÃO REQUER ATENÇÃO ({NOME_BOT})" if tipo == "ERRO" else NOME_BOT
    bot.email.enviar_email(destinatarios, assunto, html, [NOME_LOG])


def main():
    """Fluxo principal - Lançamento de Solicitações de Pagamento Geral."""
    global resultados_processamento
    x, y = pyautogui.size()
    bot.logger.informar(f"Resolução: {x} x {y}")

    resultados_processamento = {}

    processos = obter_processos()
    if not processos:
        return

    for processo in processos:
        try:
            id_tarefa = processo.id_tarefa
        except AssertionError:
            bot.logger.informar(
                f"Processo '{processo.webhook.id_processo}' não possui tarefa aberta"
            )
            processo.remover_webhook()
            continue

        status_processamento = False
        protocolo = getattr(processo.properties, "protocolo", processo.webhook.id_processo)
        filial = getattr(processo.properties, "filial", "-")

        bot.logger.informar(f"Iniciando lançamento do processo {processo.webhook.id_processo}")

        # [1] Iniciar NBS (TE02, TE03: problemas NBS/interface → retry, depois email)
        try:
            nbs = modulos.nbs.Sistema(
                *bot.configfile.obter_opcoes("nbs", ["usuario", "senha", "servidor"])
            )
            nbs.inicializar()
        except Exception as erro:
            processo.aplicar_retry(etapa="Inicialização do NBS", motivo=str(erro))
            resultados_processamento[protocolo] = [status_processamento, filial, protocolo]
            continue

        # [2] Tratar e obter dados do Holmes
        try:
            campos_obrigatorios, dados_processo = operacoes.tratar_tarefa_aberta(id_tarefa)
        except (ErroNegocio, ErroTecnico) as erro:
            bot.logger.erro(erro)
            processo.encaminhar_tarefa_erro(str(erro))
            resultados_processamento[protocolo] = [status_processamento, filial, protocolo]
            continue
        except Exception as erro:
            processo.aplicar_retry(etapa="Tratamento de dados", motivo=str(erro))
            resultados_processamento[protocolo] = [status_processamento, filial, protocolo]
            continue

        # [3] Obter Código UAB Contábil (TE07: tipo não parametrizado → não registrar, encaminhar erro)
        if not campos_obrigatorios:
            bot.logger.informar(f"Não foi possível obter os dados do processo {protocolo}")
            continue
        bot.logger.informar(f"Dados do processo extraídos para protocolo {protocolo}")

        tipo_pagamento = campos_obrigatorios[9]
        codigo_uab = modulos.nbs.solicitacao_pagamento.obter_codigo_uab_contabil(tipo_pagamento)
        if not codigo_uab:
            msg_te07 = f"Tipo de Pagamento '{tipo_pagamento}' não parametrizado na planilha TIPOS_DE_PAGAMENTO."
            bot.logger.erro(msg_te07)
            processo.encaminhar_tarefa_erro(msg_te07)
            resultados_processamento[protocolo] = [status_processamento, filial, protocolo]
            try:
                fiscais_raw = bot.configfile.obter_opcao_ou("email.destinatarios", "fiscal")
                if fiscais_raw:
                    fiscais = str(fiscais_raw).split(", ")
                    with open("email.html", encoding="utf-8") as arq:
                        html_fiscal = (
                            arq.read()
                            .replace("{0}", f"{NOME_BOT} - TE07 Tipo Pagamento não parametrizado")
                            .replace("{1}", f"{msg_te07} Não lançar no NBS.")
                            .replace("{2}", NOME_BOT)
                            .replace("{table_resultados}", "")
                        )
                    bot.email.enviar_email(fiscais, f"TE07 - {NOME_BOT}: Tipo Pagamento não parametrizado", html_fiscal, [NOME_LOG])
            except Exception as e:
                bot.logger.erro(f"Falha ao enviar email TE07 para Fiscal: {e}")
            continue

        # Fluxo Prévio II: inserir dados Holmes + Codigo_UAB_Contabil no banco de controle
        dados_extraidos = dados_processo.get("dados_extraidos")
        if dados_extraidos:
            operacoes.registrar_processo_controle(
                processo.webhook.id_processo, dados_extraidos, codigo_uab
            )

        # [4] Processar entrada no NBS (Solicitação Pagamento Geral)
        try:
            parar_apos_obs = str(bot.configfile.obter_opcao_ou("nbs", "parar_apos_observacoes") or "").lower() in ("1", "true", "sim", "s")
            entrada = modulos.nbs.processar_entrada_solicitacao_pagamento(
                campos_obrigatorios, dados_processo, parar_apos_observacoes=parar_apos_obs,
                codigo_uab_contabil=codigo_uab,
            )
        except Exception as erro:
            processo.aplicar_retry(
                etapa="Processamento da entrada no NBS", motivo=str(erro)
            )
            resultados_processamento[protocolo] = [status_processamento, filial, protocolo]
            continue

        # [5] Validar resultado
        try:
            if not entrada.SUCESSO:
                bot.logger.informar(entrada.MENSAGEM)
                processo.encaminhar_tarefa_erro(str(entrada.MENSAGEM))
                resultados_processamento[protocolo] = [status_processamento, filial, protocolo]
                continue

            processo.encaminhar_tarefa_sucesso()
            status_processamento = True
            resultados_processamento[protocolo] = [status_processamento, filial, protocolo]
        except Exception:
            continue

        bot.logger.informar(f"Processamento concluído com sucesso [{protocolo}]")


if __name__ == "__main__":
    resultados_processamento = {}
    try:
        main()
        if resultados_processamento:
            notificar_email("SUCESSO", resultados_processamento)
    except Exception:
        bot.logger.erro("Erro inesperado no fluxo")
        notificar_email("ERRO")
