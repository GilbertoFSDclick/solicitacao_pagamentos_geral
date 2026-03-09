import datetime, functools, typing
from dataclasses import dataclass, fields
import bot
from modulos.webhook import modelos, holmes

class ErroWebhook(Exception):
    def __init__ (self, mensagem: str) -> None:
        self.mensagem = mensagem
        super().__init__(self.mensagem)

@functools.cache
def client_singleton () -> bot.http.Client:
    """Criar o http `Client` configurado com o `base_url`, `token` e timeout
    - O Client ficará aberto após a primeira chamada na função devido ao `@cache`"""
    base_url, token = bot.configfile.obter_opcoes("webhook", ["base_url", "token"])
    return bot.http.Client(
        base_url = base_url,
        params = { "token": token },
        timeout = 10
    )

def checar_conexao_webhook () -> None:
    """Checar se a conexão está online
    - `AssertionError` caso negativo"""
    mensagem = "Conexão com o Webhook não detectada"
    assert client_singleton().head("/webhook/holmes").status_code == 200, mensagem

class ProcessoWebhook[T]:

    webhook: modelos.Webhook
    """Propriedades do `webhook`"""
    properties: T
    """Propriedade `processo.properties`"""
    author: modelos.Author
    """Propriedade `processo.author`"""
    documents: list[modelos.Document]
    """Propriedade `processo.documents`"""

    def __repr__ (self) -> str:
        return f"<{self.__class__.__name__} holmes({self.webhook.id_processo}) webhook({self.webhook.id_webhook})>"
    
    def remover_webhook (self) -> typing.Self:
        """Remover o processo do banco de dados do webhook"""
        bot.logger.informar(f"Removendo o {self!r}")
        response = client_singleton().delete(f"/webhook/holmes/{self.webhook.id_webhook}")
        assert response.is_success, f"Erro ao remover a {self!r} do webhook | Status code '{response.status_code}' inesperado"
        return self

    def incrementar_tentativas_webhook (self) -> typing.Self:
        """Incrementar o campo `tentativas` do processo no banco de dados do webhook
        - Incrementado `self.webhook.tentativas` para refletir com o banco"""
        bot.logger.informar(f"Incrementando o campo tentativas do {self!r}")
        response = client_singleton().patch(f"/webhook/holmes/{self.webhook.id_webhook}/tentativas")
        assert response.is_success, f"Erro ao incrementar tentativas do {self!r} no webhook | Status code '{response.status_code}' inesperado"
        self.webhook.tentativas += 1
        return self

    def atualizar_controle_webhook (self) -> typing.Self:
        """Atualizar o campo `controle` do processo no banco de dados do webhook
        - Utilizado o campo `self.webhook.controle`"""
        bot.logger.informar(f"Atualizando o campo controle {self!r}")
        response = client_singleton().put(
            f"/webhook/holmes/{self.webhook.id_webhook}/controle",
            json = self.webhook.controle
        )
        assert response.is_success, f"Erro ao atualizar o campo controle do {self!r} no webhook | Status code '{response.status_code}' inesperado"
        return self

class QueryProcessosWebhook[T]:
    def __init__(self, dms: str, filtros: typing.Dict = None, classe_validadora_properties: type[T] = dict[str, typing.Any]):
        self.dms = dms.upper()
        self.query = f"properties.DMS = '{dms.upper()}'"
        self.filtros = filtros
        self.classe_validadora_properties = classe_validadora_properties

    def procurar(self, limite: int = 999999) -> list[ProcessoWebhook[T]]:
        checar_conexao_webhook()
        response = client_singleton().get(
            "/webhook/holmes",
            params = {
                "limit": limite,
                "query": self.query,
            }
        )
        assert response.status_code == 200, f"Status code diferente de 200: '{response.status_code}'"

        resultados: list[modelos.ItemWebhook] = [
            modelos.ItemWebhook(
                id_webhook = item["id_webhook"],
                id_processo = item["id_processo"],
                criado_em = item["criado_em"],
                atualizado_em = item["atualizado_em"],
                tentativas = item["tentativas"],
                controle = item["controle"],
                dados = modelos.DadosItemWebhook(
                    author=modelos.Author(
                        **item["dados"]["author"]
                    ),
                    documents=item["dados"]["documents"],
                    properties=item["dados"]["properties"]
                )
            )
            for item in response.json()["processos"]
        ]
        if not resultados:
            bot.logger.informar(f"Nenhum resultado encontrado para o DMS {self.dms}")
            return

        # Aplica filtro nos resultados
        itens_filtrados = self.filtrar_resultados(resultados) if self.filtros else resultados

        # Gera e retorna uma lista de 'ProcessoWebhook'
        return [
            parsed for item in itens_filtrados if (parsed := self.parse(item))
        ]

    def filtrar_resultados(self, resultados: list[modelos.ItemWebhook]):
        """Aplica `self.filtros` nos resultados do webhook e os retorna"""
        return [
            item
            for item in resultados
            if all(item.dados.properties.get(campo) in valor if isinstance(valor, list)
                   else item.dados.properties.get(campo) == valor
                   for campo, valor in self.filtros.items())
        ]

    def parse(self, item_webhook: modelos.ItemWebhook) -> ProcessoWebhook[T]:
        
        pw = ProcessoWebhook()
        pw.author = item_webhook.dados.author
        pw.documents = [*item_webhook.dados.documents]
        pw.webhook = modelos.Webhook(
            id_webhook=item_webhook.id_webhook,
            id_processo=item_webhook.id_processo,
            tentativas=item_webhook.tentativas,
            controle=item_webhook.controle,
            criado_em=datetime.datetime.fromisoformat(item_webhook.criado_em),
            atualizado_em=datetime.datetime.fromisoformat(item_webhook.atualizado_em)
        )

        # Normaliza properties do webhook
        props_webhook_normalizadas = {bot.util.normalizar(k):v for k, v in item_webhook.dados.properties.items()}

        # Gera set de propriedades da classe validadora
        props_classe_validadora = {f.name for f in fields(self.classe_validadora_properties)}

        # Identifica propriedades ausentes
        props_ausentes = props_classe_validadora - set(props_webhook_normalizadas.keys())
        if props_ausentes:
            # bot.logger.informar(f"props ausentes, {props_ausentes}")
            return

        # Gera dicionário com as propriedades disponíveis na classe validadora
        props_validadas = {k: v for k, v in props_webhook_normalizadas.items() if k in props_classe_validadora}

        # Desempacota propriedades do processo na classe validadora `Properties`
        pw.properties = self.classe_validadora_properties(**props_validadas)

        return pw