"""
Webhook para Lançamento de Solicitações de Pagamento Geral.
Properties mapeadas aos campos reais do Holmes (Solicitação de Pagamentos Geral).
"""
import bot
from modulos import webhook, holmes
import typing
import dataclasses
from datetime import datetime

ACAO_SUCESSO, ACAO_ERRO, ID_PENDENCIA_TAREFA = bot.configfile.obter_opcoes(
    "holmes", ["acao_sucesso", "acao_erro", "id_pendencia_tarefa"]
)
MAX_TENTATIVAS_WEBHOOK = int(bot.configfile.obter_opcao_ou("webhook", "max_tentativas"))
NOME_ACAO_SUCESSO, NOME_ACAO_ERRO = bot.configfile.obter_opcoes(
    "holmes", ["nome_acao_tarefa_sucesso", "nome_acao_tarefa_erro"]
)


@dataclasses.dataclass
class Properties:
    """
    Propriedades do processo - nomes normalizados do Holmes.
    Campos reais: Protocolo, Filial, CPF, Tipo de pagamento, Valores, DMS.
    Execute scripts/validar_webhook.py para ver as propriedades retornadas.
    """
    protocolo: str
    filial: str
    cpf: str
    tipo_pagamento: str
    valores: str
    dms: str


class ProcessoWebhook(webhook.ProcessoWebhook[Properties]):
    def __init__(self, processo: webhook.ProcessoWebhook[Properties]) -> None:
        self.__dict__.update(processo.__dict__)

    @property
    def id_tarefa(self):
        return webhook.holmes.obter_tarefa_aberta(self.webhook.id_processo)

    @property
    def id_acao(self) -> webhook.holmes.AcaoHolmes:
        tarefa = holmes.consulta_tarefa(self.id_tarefa)
        acoes: list[dict] = tarefa.get("actions", [])

        acaoSucessoID = None
        acaoErroID = None

        for acao in acoes:
            nome, _id = acao.get("name"), acao.get("id")
            if nome in [NOME_ACAO_SUCESSO, "Avançar"]:
                acaoSucessoID = _id
            elif nome == NOME_ACAO_ERRO:
                acaoErroID = _id
        assert None not in (acaoSucessoID, acaoErroID), (
            f"A tarefa ({self.id_tarefa}) não pode ser encaminhada pois não possui ações esperadas. "
            f"Ações esperadas: [{NOME_ACAO_SUCESSO}, {NOME_ACAO_ERRO}]"
        )

        return webhook.holmes.AcaoHolmes(acaoSucessoID, acaoErroID)

    def encaminhar_tarefa_sucesso(self) -> typing.Self:
        try:
            id_tarefa = webhook.holmes.obter_tarefa_aberta(self.webhook.id_processo)
            webhook.holmes.tomar_acao_tarefa(id_tarefa, self.id_acao.SUCESSO)
            self.remover_webhook()
            bot.logger.informar("Tarefa encaminhada com sucesso. (Concluída)")
        except Exception as e:
            bot.logger.informar(f"Ocorreu um erro ao encaminhar a tarefa de sucesso: ({e})")
        return self

    def encaminhar_tarefa_erro(self, motivo_erro: str) -> typing.Self:
        try:
            id_tarefa = webhook.holmes.obter_tarefa_aberta(self.webhook.id_processo)
            webhook.holmes.tomar_acao_tarefa(
                id_tarefa,
                self.id_acao.ERRO,
                [{"id": ID_PENDENCIA_TAREFA, "value": motivo_erro}],
            )
            self.remover_webhook()
            bot.logger.informar(f"Tarefa de pendência encaminhada com sucesso. ({motivo_erro})")
        except Exception as e:
            bot.logger.informar(f"Ocorreu um erro ao encaminhar a tarefa de pendência ({e})")
        return self

    def aplicar_retry(self, etapa: str, motivo: str) -> typing.Self:
        """Aplica política de retry (TE04, TE08)."""
        try:
            if self.webhook.tentativas >= MAX_TENTATIVAS_WEBHOOK:
                bot.logger.informar(
                    f"Política de retry aplicada ({MAX_TENTATIVAS_WEBHOOK} tentativas de processamento)"
                )
                self.encaminhar_tarefa_erro(
                    f"[Política de retry] - Pendência após {MAX_TENTATIVAS_WEBHOOK} tentativas. "
                    f"|| Etapa: {etapa} || Motivo: {motivo}"
                )
                return self

            self.incrementar_tentativas_webhook()
        except Exception as erro:
            bot.logger.informar(erro)
        return self


filtros: dict[str, typing.Any] = {
    "DMS": "nbs",
}


def obter_processos():
    """Obtém processos do webhook para Lançamento de Solicitações de Pagamento Geral."""
    query = webhook.QueryProcessosWebhook[Properties](
        dms="nbs", filtros=filtros, classe_validadora_properties=Properties
    )

    processos: list[ProcessoWebhook] = []
    for processo in query.procurar() or []:
        pw = ProcessoWebhook(processo)
        processos.append(pw)

    bot.logger.informar(f"{len(processos)} processos encontrados")

    processos_ordenados = sorted(processos, key=lambda p: p.webhook.criado_em)

    return processos_ordenados
