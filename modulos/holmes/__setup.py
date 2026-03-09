# interno
import bot

# std
from datetime import datetime
from typing import Literal, Tuple


class ErroHolmes (Exception):
    """Erro próprio para o sistema HOLMES"""
    def __init__ (self, mensagem: str) -> None:
        self.mensagem = mensagem
        super().__init__(self.mensagem)


def query_tarefas_abertas (_from=0, size=200, sort="due_date", order="desc") -> Tuple[list[dict], int]:
    """Consultar tarefas abertas de nome `nome_tarefa` do fluxo `id_fluxo` atribuídas ao `id_usuário`
    - Variáveis utilizadas `[holmes] -> "host", "token", "id_fluxo", "nome_tarefa", "id_usuario"`
    - Exceção `ErroHolmes` caso ocorra algum problema"""
    bot.logger.informar("Procurando por tarefas abertas")
    # obter as variáveis de ambiente
    host, token, fluxo, processo, usuario = bot.configfile.obter_opcoes("holmes", ["host", "token", "id_fluxo", "nome_processo", "id_usuario"])
    
    # realizar a chamada HTTP
    url, headers = rf"{ host }/v2/search", { "api_token": token, "Accept": "application/json" }
    body = {
        "query": {
            "context": "task",
            "from": _from,
            "size": size,
            "sort": sort,
            "order": order,
            "groups": [
                {
                    "terms": [
                        {
                            "field": "template_id",
                            "type": "is",
                            "value": str(fluxo)
                        },
                        {
                            "field": "status",
                            "type": "is",
                            "value": "opened"
                        },
                        {
                            "field": "process_name",
                            "type": "match_phrase",
                            "value": str(processo)
                        },
                        {
                            "field": "assignee_id",
                            "type": "is",
                            "value": str(usuario)
                        }
                    ]
                }
            ]
        }
    }
    response = bot.http.request("POST", url, headers=headers, json=body, timeout=120)
    
    # bot.logger.debug(f"""Resposta da chamada:
    #     codigo: { response.status_code }
    #     headers: { response.headers }
    #     body: { response.text }""")

    # validações
    if response.status_code != 200:
        raise ErroHolmes(f"O status code '{ response.status_code }' foi diferente do esperado")
    
    body = response.json()
    if not isinstance(body, dict) or not isinstance(body.get("docs"), list):
        raise ErroHolmes(f"O conteúdo json não possui propriedade esperada")

    # OK
    bot.logger.informar(f"Consulta resultou em { len(body['docs']) } tarefa(s) de um total de { body['total'] }")
    return (body["docs"], body["total"])


def consulta_tarefa (identificador: str) -> dict:
    """Consultar tarefa `identificador`
    - Variáveis utilizadas `[holmes] -> "host", "token"`
    - Exceção `ErroHolmes` caso ocorra algum problema"""
    bot.logger.informar(f"Consultando tarefa({ identificador })")
    # obter as variáveis de ambiente
    host, token = bot.configfile.obter_opcoes("holmes", ["host", "token"])
    
    # realizar a chamada HTTP
    url = rf"{ host }/v1/tasks/{ identificador }"
    headers = { "api_token": token, "Accept": "application/json" }
    response = bot.http.request("GET", url, headers=headers, timeout=120)
    
    # bot.logger.debug(f"""Resposta da chamada:
    #     codigo: { response.status_code }
    #     headers: { response.headers }
    #     body: { response.text }""")

    # validações
    if response.status_code != 200:
        raise ErroHolmes(f"O status code '{ response.status_code }' foi diferente do esperado")
    
    body = response.json()
    if not isinstance(body, dict):
        raise ErroHolmes(f"O conteúdo json não possui o formato esperado")

    # OK
    bot.logger.informar(f"Consulta realizada com sucesso")
    return body


def consulta_documento_tarefa (tarefaID: str, documentoID: str) -> tuple[str, str | None, bytes]:
    """Consultar tarefa `identificador`
    - Variáveis utilizadas `[holmes] -> "host", "token"`
    - Exceção `ErroHolmes` caso ocorra algum problema
    - Retorna `(content-type, encode, bytes)` do documento"""
    bot.logger.informar(f"Consultando documento({ documentoID }) da tarefa({ tarefaID })")
    # obter as variáveis de ambiente
    host, token = bot.configfile.obter_opcoes("holmes", ["host", "token"])
    
    # realizar a chamada HTTP
    url = rf"{ host }/v1/tasks/{ tarefaID }/documents/{ documentoID }"
    headers = { "api_token": token }
    response = bot.http.request("GET", url, headers=headers, timeout=120)
    
    bot.logger.debug(f"""Resposta da chamada:
        codigo: { response.status_code }
        headers: { response.headers }""")

    # validações
    if response.status_code != 200:
        raise ErroHolmes(f"O status code '{ response.status_code }' foi diferente do esperado")

    # OK
    bot.logger.informar(f"Consulta realizada com sucesso")
    return (response.headers.get("Content-Type", ""), 
            response.charset_encoding,
            response.content)


def consulta_processo (identificador: str) -> dict:
    """Consultar processo `identificador`
    - Variáveis utilizadas `[holmes] -> "host", "token"`
    - Exceção `ErroHolmes` caso ocorra algum problema"""
    bot.logger.informar(f"Consultando processo({ identificador })")
    # obter as variáveis de ambiente
    host, token = bot.configfile.obter_opcoes("holmes", ["host", "token"])
    
    # realizar a chamada HTTP
    url = rf"{ host }/v1/processes/{ identificador }/details"
    headers = { "api_token": token, "Accept": "application/json" }
    response = bot.http.request("GET", url, headers=headers, timeout=120)

    # bot.logger.debug(f"""Resposta da chamada:
    #     codigo: { response.status_code }
    #     headers: { response.headers }
    #     body: { response.text }""")

    # validações
    if response.status_code != 200:
        raise ErroHolmes(f"O status code '{ response.status_code }' foi diferente do esperado")
    
    body = response.json()
    if not isinstance(body, dict):
        raise ErroHolmes(f"O conteúdo json não possui o formato esperado")

    # OK
    bot.logger.informar(f"Consulta realizada com sucesso")
    return body


def encaminhar_tarefa (tarefaID: str, acaoID: str, propriedades: list[dict] = []) -> None:
    """Encaminhar tarefa `tarefaID` para a ação `acaoID`
    - Variáveis utilizadas `[holmes] -> "host", "token"`
    - Exceção `ErroHolmes` caso ocorra algum problema
    - utlizar `propriedades` caso tenha de se passar um adicional (motivo de pendência)"""
    bot.logger.informar(f"Encaminhando tarefa({ tarefaID }) para a ação({ acaoID })")
    # obter as variáveis de ambiente
    host, token = bot.configfile.obter_opcoes("holmes", ["host", "token"])

    # realizar a chamada HTTP
    url = rf"{ host }/v1/tasks/{ tarefaID }/action"
    headers = { "api_token": token, "Accept": "application/json" }
    body = {
        "task": { 
            "action_id": acaoID,
            "confirm_action": True,
            "property_values": propriedades
        }
    }
    response = bot.http.request("POST", url, headers=headers, json=body, timeout=120)

    # bot.logger.debug(f"""Resposta da chamada:
    #     codigo: { response.status_code }
    #     headers: { response.headers }
    #     body: { response.text }""")

    # validações
    if response.status_code != 200:
        raise ErroHolmes(f"O status code '{ response.status_code }' foi diferente do esperado")

    # OK
    bot.logger.informar(f"Tarefa encaminhada com sucesso")

def filtrar_tarefas (tarefas: list[str],
                     filtros: dict[str, str | list],
                     operacao: Literal["ALL", "ANY"] = "ALL",
                     ordenar_prioridade_vencimento: bool = True) -> list[tuple[str, ...]]:
    """Retorna todas as tarefas filtradas com base nos parâmetros \n
    - `tarefas` lista de ID das tarefas obtidas na query principal \n
    - `filtros = {"propriedade": "valor"}` são os parâmetros de filtragem \n
    - `operacao` define a combinação de correspondência para múltiplas regras (AND / OR)
    - `ordenar_prioridade_vencimento` define se a lista de tarefas filtradas deve ser ordenada por data de vencimento (vencimento mais próximo vem primeiro)"""

    NOME_ACAO_SUCESSO, NOME_ACAO_ERRO, ID_PENDENCIA = bot.configfile.obter_opcoes("holmes", ["nome_acao_tarefa_sucesso", "nome_acao_tarefa_erro", "id_pendencia_tarefa"])

    bot.logger.informar(f"Filtrando tarefas com os parâmetros fornecidos")
    tarefasFiltradas = []

    for identificador in tarefas:
        tarefa = consulta_tarefa(identificador)
        propriedades: list[dict] = tarefa.get("properties", [])
        resultado_filtros = []
        torre = None
        acaoErroID: str = None
        acaoSucessoID: str = None
        data_vencimento = datetime.now().replace(microsecond=0).isoformat() + ".000Z" # Formato de data retornada no Holmes (ISO 8601)
        
        # Obter ID das ações da tarefa
        try:
            acoes: list[dict] = tarefa.get("actions", [])

            # encontrar ações obrigatórias
            for acao in acoes:
                nome, _id = acao.get("name"), acao.get("id")
                if nome == NOME_ACAO_SUCESSO: acaoSucessoID = _id
                elif nome == NOME_ACAO_ERRO: acaoErroID = _id

            assert None not in (acaoSucessoID, acaoErroID), f"A tarefa ({ identificador }) não pode ser encaminhada pois não possui ações esperadas\n\tAções esperadas: [{ NOME_ACAO_SUCESSO }, { NOME_ACAO_ERRO }]"
        except Exception as erro:
            bot.logger.informar( erro )
            continue
        
        # Extrair torre
        for propriedade in propriedades:
            nome: str = propriedade.get("name", "")
            if nome.lower() == "torre":
                torre: str = propriedade.get("value", "")

        # Extrair data de vencimento
        for propriedade in propriedades:
            nome: str = propriedade.get("name", "")
            if nome.lower() == "data de vencimento":
                data_vencimento: str = propriedade.get("value", "")

        if not torre:
            # Pendencia se não tiver o campo Torre
            mensagem = "Torre não identificada"
            propriedades = [{ "id": ID_PENDENCIA, "value": mensagem}]
            encaminhar_tarefa(identificador, acaoErroID, propriedades)
            continue
        
        # Aplicar filtros
        for nome_propriedade, valor_propriedade in filtros.items():
            filtro_correspondido = False
            if isinstance(valor_propriedade, str):
                for propriedade in propriedades:
                    nome: str = propriedade.get("name", "")
                    valor: str | list[dict] = propriedade.get("value", "")
                    if valor is None: continue
                    
                    if isinstance(valor, list):
                        for item in valor:
                            item: dict
                            valor: str = item.get("text", "")

                    if nome.lower() == nome_propriedade.lower() and valor.lower() == valor_propriedade.lower():
                        filtro_correspondido = True
                        break

            elif isinstance(valor_propriedade, list):
                for item_propriedade in valor_propriedade:
                    item_propriedade: str
                    for propriedade in propriedades:
                        nome: str = propriedade.get("name", "")
                        valor: str | list[dict] = propriedade.get("value")
                        if valor is None: continue
                        
                        if isinstance(valor, list):
                            for item in valor:
                                item: dict
                                valor: str = item.get("text", "")

                        if nome.lower() == nome_propriedade.lower() and valor.lower() == item_propriedade.lower():
                            filtro_correspondido = True
                            break
            
            resultado_filtros.append(filtro_correspondido)

        if operacao == "ALL" and all(resultado_filtros):
            tarefasFiltradas.append( (identificador, torre, data_vencimento) ) # Data de vencimento não é considerada no retorno da filtragem, é usada apenas para ordenação
        elif operacao == "ANY" and any(resultado_filtros):
            tarefasFiltradas.append( (identificador, torre, data_vencimento) )

    if ordenar_prioridade_vencimento:
        # Filtrar fora as tarefas sem data de vencimento
        tarefas_validas = [t for t in tarefasFiltradas if t[2] is not None]

        # Ordenar somente as válidas
        tarefas_validas = sorted(
            tarefas_validas,
            key=lambda x: datetime.fromisoformat(str(x[2]).replace("Z", "+00:00"))
        )

        # Manter apenas identificador e torre
        tarefasFiltradas = [(identificador, torre) for identificador, torre, _ in tarefas_validas]

        bot.logger.informar(
            f"""Filtragem resultou em {len(tarefasFiltradas)} tarefa(s) 
            {tarefasFiltradas}"""
        )

    return tarefasFiltradas

__all__ = [
    "ErroHolmes",
    "consulta_tarefa",
    "consulta_processo",
    "encaminhar_tarefa",
    "query_tarefas_abertas",
    "consulta_documento_tarefa",
    "filtrar_tarefas"
]
