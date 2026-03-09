# std
import base64, certifi, mimetypes
from functools import cache
from typing import Callable, Generator, Literal, Self, NamedTuple
# externo
import bot

class AcaoHolmes(NamedTuple):
    SUCESSO: str
    ERRO: str

@cache
def client_singleton () -> bot.http.Client:
    """Criar o http `Client` configurado com o `host`, `token` e timeout
    - O Client ficará aberto após a primeira chamada na função devido ao `@cache`"""
    host, token = bot.configfile.obter_opcoes("holmes", ["host", "token"])
    return bot.http.Client(
        base_url = host,
        headers  = { "api_token": token },
        timeout  = 120,
        verify   = certifi.where(),
    )

def obter_tarefa_aberta (id_processo: str):
    """Consultar o processo `id_processo`
    - Variáveis utilizadas `[holmes] -> "host", "token"`"""
    bot.logger.informar(f"Consultando processo({id_processo}) no Holmes")

    nome_atividade: str = bot.configfile.obter_opcao_ou("holmes", "nome_atividade")

    response = client_singleton().get(f"/v1/processes/{id_processo}")
    assert response.status_code == 200, f"Status code '{response.status_code}' diferente do esperado"
    
    current_activities: list[dict] = response.json()['current_activities']
    tarefa_aberta: str | None = next(
        (tarefa['id'] for tarefa in current_activities
         if tarefa['status'] == 'opened' and tarefa['name'] == nome_atividade), None
    )
    
    assert tarefa_aberta, "Não há tarefa aberta para este processo"
    return tarefa_aberta

def tomar_acao_tarefa (
        id_tarefa: str,
        id_acao: str,
        propriedades: list[dict[Literal["id", "value", "text"], str]] | None = None
    ) -> None:
    """Tomar `acao` na `tarefa`
    - `propriedades` caso seja necessário informar algum adicional (motivo de pendência)
    - Variáveis utilizadas `[holmes] -> "host", "token"`"""
    bot.logger.informar(f"Tomando ação({id_acao}) na tarefa({id_tarefa}) no Holmes")

    response = client_singleton().post(
        f"/v1/tasks/{id_tarefa}/action",
        json = {
            "task": { 
                "action_id": id_acao,
                "confirm_action": True,
                "property_values": propriedades or []
            }
    })
    assert response.status_code == 200, f"Status code '{response.status_code}' diferente do esperado ao tomar ação em tarefa no Holmes"
    
