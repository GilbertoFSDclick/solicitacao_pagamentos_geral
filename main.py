"""
Fluxo principal - Lançamento de Solicitações de Pagamento Geral.
Conforme Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral.
"""
import bot
import modulos
import operacoes
import pyautogui
from pathlib import Path
from time import sleep
from types import SimpleNamespace
from src.webhook import obter_processos
from src.exceptions import ErroNegocio, ErroTecnico

from typing import Literal

NOME_BOT = bot.configfile.obter_opcao_ou("bot", "nome")
NOME_LOG = ".log"


def _flag_true(valor) -> bool:
    return str(valor or "").strip().lower() in ("1", "true", "sim", "s", "yes", "y", "on")


def _modo_teste_sem_perda() -> bool:
    """Modo seguro: não remove webhook, não encaminha tarefa e não aplica retry no webhook."""
    return _flag_true(bot.configfile.obter_opcao_ou("bot", "modo_teste_sem_perda"))


def _parar_antes_confirmar() -> bool:
    """Se True, para ANTES de clicar Confirmar no NBS (ideal para testes sem executar)."""
    return _flag_true(bot.configfile.obter_opcao_ou("bot", "parar_antes_confirmar"))


def _ids_teste_processos() -> list[str]:
    """Lê lista de IDs de processo para teste no formato CSV em [bot] ids_teste_processos."""
    raw = bot.configfile.obter_opcao_ou("bot", "ids_teste_processos") or ""
    return [p.strip() for p in str(raw).split(",") if p and p.strip()]


class _ProcessoTesteSemWebhook:
    """Adapter mínimo para executar fluxo por ID de processo sem depender do webhook."""

    def __init__(self, id_processo: str):
        self.webhook = SimpleNamespace(id_processo=id_processo, tentativas=0)
        self.properties = SimpleNamespace(protocolo=id_processo, filial="-")
        self._id_tarefa_cache = None

    @property
    def id_tarefa(self):
        if not self._id_tarefa_cache:
            self._id_tarefa_cache = modulos.webhook.holmes.obter_tarefa_aberta(
                self.webhook.id_processo
            )
        return self._id_tarefa_cache

    def remover_webhook(self):
        return self

    def encaminhar_tarefa_sucesso(self):
        return self

    def encaminhar_tarefa_erro(self, _motivo: str):
        return self

    def aplicar_retry(self, etapa: str, motivo: str):
        bot.logger.informar(
            f"[MODO TESTE SEM WEBHOOK] Retry simulado | etapa={etapa} | motivo={motivo}"
        )
        return self

    @property
    def retry_esgotado(self) -> bool:
        return False


def _obter_anexos_log() -> list[str]:
    """Retorna anexos de log existentes para envio de e-mail.
    Evita falha por caminho fixo inexistente."""
    candidatos: list[str] = []
    log_cfg = bot.configfile.obter_opcao_ou("logger", "arquivo")
    if log_cfg:
        candidatos.append(str(log_cfg).strip())

    candidatos.extend([
        NOME_LOG,
        f"{NOME_BOT}.log",
        "bot.log",
        "logs/bot.log",
        "logs/.log",
    ])

    anexos: list[str] = []
    vistos: set[str] = set()
    for candidato in candidatos:
        caminho = str(candidato).strip()
        if not caminho or caminho in vistos:
            continue
        vistos.add(caminho)
        try:
            if bot.windows.afirmar_arquivo(caminho) or Path(caminho).exists():
                anexos.append(caminho)
                break
        except Exception:
            continue

    if not anexos:
        bot.logger.informar("Nenhum arquivo de log encontrado para anexar no e-mail")
    return anexos


def _obter_destinatarios_email(chave: str) -> list[str]:
    """Obtém lista de destinatários a partir de [email.destinatarios]."""
    try:
        raw = bot.configfile.obter_opcao_ou("email.destinatarios", chave) or ""
    except Exception:
        raw = ""
    return [x.strip() for x in str(raw).split(",") if x and x.strip()]


def _obter_anexos_execucao_sucesso() -> list[str]:
    """Monta anexos do e-mail de sucesso: log + anexos opcionais da documentação."""
    anexos = list(_obter_anexos_log())

    # Opcional: permitir anexos adicionais via .ini sem tornar obrigatório.
    candidatos_cfg = []
    for chave in ("relatorio_final", "planilha_atualizada", "notas_fiscais_compactadas"):
        try:
            valor = bot.configfile.obter_opcao_ou("email.anexos", chave)
            if valor:
                candidatos_cfg.append(str(valor).strip())
        except Exception:
            continue

    vistos = set(anexos)
    for candidato in candidatos_cfg:
        if not candidato or candidato in vistos:
            continue
        try:
            if bot.windows.afirmar_arquivo(candidato) or Path(candidato).exists():
                anexos.append(candidato)
                vistos.add(candidato)
        except Exception:
            continue
    return anexos


def _obter_processos_com_retry() -> list:
    """Consulta webhook com retry para reduzir falhas transitórias de rede/serviço."""
    try:
        max_tentativas = int(bot.configfile.obter_opcao_ou("webhook", "max_tentativas") or 3)
    except Exception:
        max_tentativas = 3
    max_tentativas = max(1, min(max_tentativas, 5))

    ultimo_erro: Exception | None = None
    for tentativa in range(1, max_tentativas + 1):
        try:
            processos = obter_processos() or []
            if tentativa > 1:
                bot.logger.informar(f"Webhook respondeu na tentativa {tentativa}/{max_tentativas}")
            return processos
        except Exception as erro:
            ultimo_erro = erro
            bot.logger.erro(
                f"Falha ao obter processos do webhook (tentativa {tentativa}/{max_tentativas}): {erro}"
            )
            if tentativa < max_tentativas:
                sleep(2 * tentativa)

    raise ErroTecnico(f"Falha ao consultar webhook após {max_tentativas} tentativas: {ultimo_erro}")


def _obter_processos_teste_por_ids(ids_processos: list[str]) -> list:
    """Monta lista de processos de teste a partir de IDs Holmes, sem webhook."""
    processos: list[_ProcessoTesteSemWebhook] = []
    for id_processo in ids_processos:
        processos.append(_ProcessoTesteSemWebhook(id_processo))
    bot.logger.informar(
        f"[MODO TESTE SEM WEBHOOK] {len(processos)} processo(s) configurado(s) por IDs"
    )
    return processos


def _notificar_email_automatico(
    tipo: Literal["ERRO", "SUCESSO"],
    resultados: dict | None = None,
    contexto: str = "",
) -> None:
    """Envia e-mail automático apenas fora do modo de teste sem perda."""
    if _modo_teste_sem_perda():
        sufixo = f" | contexto={contexto}" if contexto else ""
        bot.logger.informar(
            f"[MODO TESTE SEM PERDA] E-mail automático suprimido (tipo={tipo}){sufixo}"
        )
        return
    notificar_email(tipo, resultados)


def _classificar_excecao_operacional(mensagem: str) -> str | None:
    """Classifica exceções operacionais relevantes para tratamento simplificado."""
    msg = str(mensagem or "").lower()
    termos_te06 = (
        "out of memory",
        "insufficient memory",
        "memoria",
        "memória",
        "sem memória",
        "no space left",
        "disk full",
        "espaço em disco",
        "espaco em disco",
    )
    if any(t in msg for t in termos_te06):
        return "TE06"
    return None


def _tentar_alocar_manual_te05(processo, motivo: str) -> None:
    """TE05: tenta alocar tarefa para responsável manual após pendência de processamento."""
    try:
        id_tarefa = processo.id_tarefa
        modulos.webhook.holmes.alocar_tarefa_manual(id_tarefa, f"TE05 - {motivo}")
    except Exception as e_aloc:
        bot.logger.informar(f"TE05: falha ao alocar tarefa para manual: {e_aloc}")


def notificar_email(tipo: Literal["ERRO", "SUCESSO"], resultados: dict | None = None) -> None:
    """Enviar notificação via e-mail (TE02, TE03, TE04, TE07)."""
    html: str
    mensagem = (
        "Bot executado com sucesso"
        if tipo == "SUCESSO"
        else "A execução da automação resultou em erro"
    )
    destinatarios = _obter_destinatarios_email(str(tipo).lower())
    if not destinatarios:
        bot.logger.informar(f"Sem destinatários configurados para e-mail de {tipo}")
        return

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
    anexos = _obter_anexos_execucao_sucesso() if tipo == "SUCESSO" else _obter_anexos_log()
    bot.email.enviar_email(destinatarios, assunto, html, anexos)


def _processar_um_processo(processo, modo_teste_sem_perda: bool = False) -> tuple[bool, str, str]:
    """
    Processa um único processo do webhook.
    Retorna (status_processamento, filial, protocolo).
    """
    status_processamento = False
    protocolo = getattr(processo.properties, "protocolo", processo.webhook.id_processo)
    filial = getattr(processo.properties, "filial", "-")

    bot.logger.informar(
        f"Iniciando lançamento do processo {processo.webhook.id_processo}"
        + (" [MODO TESTE SEM PERDA]" if modo_teste_sem_perda else "")
    )

    msg_erro = None
    
    try:
        # [1] Iniciar NBS (TE02, TE03: problemas NBS/interface → retry, depois email)
        try:
            nbs = modulos.nbs.Sistema(
                *bot.configfile.obter_opcoes("nbs", ["usuario", "senha", "servidor"])
            )
            nbs.inicializar()
        except Exception as erro:
            msg_erro = f"Inicialização do NBS: {str(erro)}"
            if _classificar_excecao_operacional(msg_erro) == "TE06":
                msg_te06 = f"TE06 - Erro de memória/disco na inicialização do NBS: {erro}"
                bot.logger.erro(msg_te06)
                if not modo_teste_sem_perda:
                    processo.encaminhar_tarefa_erro(msg_te06)
                else:
                    bot.logger.informar("[MODO TESTE SEM PERDA] TE06 detectado sem encaminhar tarefa")
                return status_processamento, filial, protocolo
            if not modo_teste_sem_perda:
                processo.aplicar_retry(etapa="Inicialização do NBS", motivo=str(erro))
                if processo.retry_esgotado:
                    _notificar_email_automatico(
                        "ERRO",
                        {protocolo: [False, filial, protocolo]},
                        contexto="retry_esgotado_inicializacao_nbs",
                    )
            return status_processamento, filial, protocolo

        # [2] Tratar e obter dados do Holmes
        try:
            campos, dados_processo = operacoes.tratar_tarefa_aberta(
                processo.id_tarefa
            )
        except ErroNegocio as erro:
            bot.logger.erro(str(erro))
            msg_erro = str(erro)
            if not modo_teste_sem_perda:
                processo.encaminhar_tarefa_erro(str(erro))
            else:
                bot.logger.informar("[MODO TESTE SEM PERDA] Erro negócio - não encaminhando")
            return status_processamento, filial, protocolo
        except ErroTecnico as erro:
            # TE08/TE04: falha técnica no Holmes deve seguir política de retry.
            msg_erro = f"Extração de dados Holmes: {str(erro)}"
            if not modo_teste_sem_perda:
                processo.aplicar_retry(etapa="Extração de dados Holmes", motivo=str(erro))
                if processo.retry_esgotado:
                    _notificar_email_automatico(
                        "ERRO",
                        {protocolo: [False, filial, protocolo]},
                        contexto="retry_esgotado_extracao_holmes",
                    )
            return status_processamento, filial, protocolo
        except Exception as erro:
            msg_erro = f"Tratamento de dados: {str(erro)}"
            if _classificar_excecao_operacional(msg_erro) == "TE06":
                msg_te06 = f"TE06 - Erro de memória/disco no tratamento de dados: {erro}"
                bot.logger.erro(msg_te06)
                if not modo_teste_sem_perda:
                    processo.encaminhar_tarefa_erro(msg_te06)
                else:
                    bot.logger.informar("[MODO TESTE SEM PERDA] TE06 detectado sem encaminhar tarefa")
                return status_processamento, filial, protocolo
            if not modo_teste_sem_perda:
                processo.aplicar_retry(etapa="Tratamento de dados", motivo=str(erro))
                if processo.retry_esgotado:
                    _notificar_email_automatico(
                        "ERRO",
                        {protocolo: [False, filial, protocolo]},
                        contexto="retry_esgotado_tratamento_dados",
                    )
            return status_processamento, filial, protocolo

        # [3] Obter Código UAB Contábil (TE07: tipo não parametrizado → não registrar, encaminhar erro)
        bot.logger.informar(
            f"Dados do processo extraídos para protocolo {protocolo}"
        )

        tipo_pagamento = campos.tipo_pagamento
        codigo_uab = modulos.nbs.solicitacao_pagamento.obter_codigo_uab_contabil(
            tipo_pagamento
        )
        if not codigo_uab:
            msg_te07 = (
                f"Tipo de Pagamento '{tipo_pagamento}' não parametrizado na planilha "
                "TIPOS_DE_PAGAMENTO."
            )
            msg_erro = msg_te07
            bot.logger.erro(msg_te07)
            if not modo_teste_sem_perda:
                processo.encaminhar_tarefa_erro(msg_te07)
            else:
                bot.logger.informar("[MODO TESTE SEM PERDA] TE07 detectado sem encaminhar tarefa")
            if modo_teste_sem_perda:
                bot.logger.informar(
                    "[MODO TESTE SEM PERDA] E-mail TE07 para Fiscal suprimido"
                )
            else:
                try:
                    fiscais = _obter_destinatarios_email("fiscal")
                    if fiscais:
                        with open("email.html", encoding="utf-8") as arq:
                            html_fiscal = (
                                arq.read()
                                .replace(
                                    "{0}",
                                    f"{NOME_BOT} - TE07 Tipo Pagamento não parametrizado",
                                )
                                .replace("{1}", f"{msg_te07} Não lançar no NBS.")
                                .replace("{2}", NOME_BOT)
                                .replace("{table_resultados}", "")
                            )
                        bot.email.enviar_email(
                            fiscais,
                            f"TE07 - {NOME_BOT}: Tipo Pagamento não parametrizado",
                            html_fiscal,
                            _obter_anexos_log(),
                        )
                except Exception as e:
                    bot.logger.erro(f"Falha ao enviar email TE07 para Fiscal: {e}")
            return status_processamento, filial, protocolo

        # Fluxo Prévio II: inserir dados Holmes + Codigo_UAB_Contabil no banco de controle
        dados_extraidos = dados_processo.get("dados_extraidos")
        if dados_extraidos:
            operacoes.registrar_processo_controle(
                processo.webhook.id_processo, dados_extraidos, codigo_uab
            )

        # [4] Processar entrada no NBS (Solicitação Pagamento Geral)
        try:
            parar_apos_obs = (
                str(
                    bot.configfile.obter_opcao_ou(
                        "nbs", "parar_apos_observacoes"
                    )
                    or ""
                ).lower()
                in ("1", "true", "sim", "s")
            )
            entrada = modulos.nbs.processar_entrada_solicitacao_pagamento(
                campos,
                dados_processo,
                parar_apos_observacoes=parar_apos_obs,
                parar_antes_confirmar=_parar_antes_confirmar(),
                codigo_uab_contabil=codigo_uab,
            )
        except Exception as erro:
            msg_erro = f"Processamento da entrada no NBS: {str(erro)}"
            bot.logger.erro(msg_erro)
            if _classificar_excecao_operacional(msg_erro) == "TE06":
                msg_te06 = f"TE06 - Erro de memória/disco no processamento do NBS: {erro}"
                bot.logger.erro(msg_te06)
                if not modo_teste_sem_perda:
                    processo.encaminhar_tarefa_erro(msg_te06)
                else:
                    bot.logger.informar("[MODO TESTE SEM PERDA] TE06 detectado sem encaminhar tarefa")
                return status_processamento, filial, protocolo
            if not modo_teste_sem_perda:
                processo.aplicar_retry(
                    etapa="Processamento da entrada no NBS", motivo=str(erro)
                )
                if processo.retry_esgotado:
                    _notificar_email_automatico(
                        "ERRO",
                        {protocolo: [False, filial, protocolo]},
                        contexto="retry_esgotado_processamento_nbs",
                    )
            return status_processamento, filial, protocolo

        # [5] Validar resultado
        try:
            if not entrada.SUCESSO:
                msg_erro = str(entrada.MENSAGEM)
                bot.logger.informar(msg_erro)
                if not modo_teste_sem_perda:
                    processo.encaminhar_tarefa_erro(str(entrada.MENSAGEM))
                    _tentar_alocar_manual_te05(processo, str(entrada.MENSAGEM))
                else:
                    bot.logger.informar("[MODO TESTE SEM PERDA] Resultado com erro sem encaminhar tarefa")
                return status_processamento, filial, protocolo

            if not modo_teste_sem_perda:
                processo.encaminhar_tarefa_sucesso()
            else:
                bot.logger.informar("[MODO TESTE SEM PERDA] Fluxo OK sem concluir/remover tarefa")
            status_processamento = True
        except Exception as erro_enc:
            # Se falhar na etapa de encaminhamento, apenas registra e segue
            msg_erro = f"Encaminhamento de tarefa: {str(erro_enc)}"
            bot.logger.informar(msg_erro)

        bot.logger.informar(f"Processamento concluído com sucesso [{protocolo}]")
    
    except Exception as erro_geral:
        # CATCH-ALL: Se qualquer exceção inesperada ocorrer
        msg_erro = f"ERRO INESPERADO: {str(erro_geral)}"
        bot.logger.erro(msg_erro)
        if not modo_teste_sem_perda:
            try:
                processo.encaminhar_tarefa_erro(msg_erro)
            except Exception:
                bot.logger.erro(f"Falha ao encaminhar erro para Holmes: {msg_erro}")
    
    return status_processamento, filial, protocolo


def main():
    """Fluxo principal - Lançamento de Solicitações de Pagamento Geral."""
    global resultados_processamento
    x, y = pyautogui.size()
    bot.logger.informar(f"Resolução: {x} x {y}")

    resultados_processamento = {}
    modo_teste_sem_perda = _modo_teste_sem_perda()
    parar_antes_confirmar = _parar_antes_confirmar()
    ids_teste = _ids_teste_processos()
    if modo_teste_sem_perda:
        bot.logger.informar(
            "MODO TESTE SEM PERDA ativo: não encaminha tarefa no Holmes, "
            "não remove webhook. Executa o fluxo NBS até o final."
        )
    if parar_antes_confirmar:
        bot.logger.informar(
            "MODO PARAR ANTES DE CONFIRMAR ativo: para antes de clicar Confirmar no NBS."
        )
    if ids_teste:
        bot.logger.informar(
            f"MODO TESTE POR IDs ativo: {ids_teste}"
        )

    processos = []
    if modo_teste_sem_perda and ids_teste:
        processos = _obter_processos_teste_por_ids(ids_teste)
    else:
        try:
            processos = _obter_processos_com_retry()
        except Exception as erro:
            bot.logger.erro(f"Não foi possível obter processos do webhook: {erro}")
            _notificar_email_automatico("ERRO", contexto="falha_obter_processos_webhook")
            return

    if not processos:
        return

    for processo in processos:
        try:
            # Gatilho: garante que a tarefa está aberta
            _ = processo.id_tarefa
        except AssertionError as erro:
            bot.logger.informar(
                f"Processo '{processo.webhook.id_processo}' sem tarefa aberta/consulta inválida: {erro}"
            )
            if not modo_teste_sem_perda:
                processo.remover_webhook()
            else:
                bot.logger.informar("[MODO TESTE SEM PERDA] Webhook preservado")
            continue
        except Exception as erro:
            bot.logger.erro(
                f"Falha ao validar tarefa aberta para processo '{processo.webhook.id_processo}': {erro}"
            )
            if not modo_teste_sem_perda:
                processo.aplicar_retry(etapa="Obter tarefa aberta Holmes", motivo=str(erro))
            continue

        status_processamento, filial, protocolo = _processar_um_processo(
            processo,
            modo_teste_sem_perda=modo_teste_sem_perda,
        )
        resultados_processamento[protocolo] = [
            status_processamento,
            filial,
            protocolo,
        ]


if __name__ == "__main__":
    resultados_processamento = {}
    try:
        main()
        if resultados_processamento:
            resultados_sucesso = {
                k: v for k, v in resultados_processamento.items() if bool(v[0])
            }
            resultados_erro = {
                k: v for k, v in resultados_processamento.items() if not bool(v[0])
            }
            if resultados_sucesso:
                _notificar_email_automatico(
                    "SUCESSO",
                    resultados_sucesso,
                    contexto="resumo_execucao",
                )
            if resultados_erro:
                _notificar_email_automatico(
                    "ERRO",
                    resultados_erro,
                    contexto="resumo_execucao",
                )
    except Exception:
        bot.logger.erro("Erro inesperado no fluxo")
        _notificar_email_automatico("ERRO", contexto="erro_inesperado_main")
