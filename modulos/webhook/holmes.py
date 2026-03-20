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
    
    current_activities: list[dict] = response.json().get('current_activities', [])

    # 1) Preferência por atividade aberta com nome configurado (match exato)
    tarefa_aberta: str | None = next(
        (
            tarefa.get('id') for tarefa in current_activities
            if tarefa.get('status') == 'opened' and tarefa.get('name') == nome_atividade
        ),
        None,
    )

    # 2) Fallback: match normalizado/contido (evita falha por pequena variação de nome)
    if not tarefa_aberta and nome_atividade:
        nome_cfg_norm = bot.util.normalizar(str(nome_atividade))
        for tarefa in current_activities:
            if tarefa.get('status') != 'opened':
                continue
            nome_tarefa = str(tarefa.get('name') or "")
            nome_tarefa_norm = bot.util.normalizar(nome_tarefa)
            if nome_cfg_norm and (
                nome_tarefa_norm == nome_cfg_norm
                or nome_cfg_norm in nome_tarefa_norm
                or nome_tarefa_norm in nome_cfg_norm
            ):
                tarefa_aberta = tarefa.get('id')
                break

    # 3) Fallback final: primeira atividade aberta
    if not tarefa_aberta:
        tarefa_aberta = next(
            (tarefa.get('id') for tarefa in current_activities if tarefa.get('status') == 'opened'),
            None,
        )
        if tarefa_aberta:
            nomes_abertos = [
                str(t.get('name') or '')
                for t in current_activities
                if t.get('status') == 'opened'
            ]
            bot.logger.informar(
                f"Fallback tarefa aberta aplicado para processo({id_processo}). "
                f"nome_atividade_config='{nome_atividade}' | abertas={nomes_abertos}"
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


def alocar_tarefa_manual(id_tarefa: str, mensagem: str = "Lançamento Manual RPA falhou") -> None:
    """Aloca a tarefa para o usuário manual configurado em `[holmes] id_usuario_manual`.
    Documentação TE04/TE05: alocar 'Lançamento Manual RPA falhou'.
    Se `id_usuario_manual` não estiver configurado, apenas loga e ignora."""
    id_usuario = bot.configfile.obter_opcao_ou("holmes", "id_usuario_manual")
    if not id_usuario:
        bot.logger.informar(
            "[Holmes] id_usuario_manual não configurado — alocação manual ignorada. "
            f"Mensagem: {mensagem}"
        )
        return

    bot.logger.informar(
        f"Alocando tarefa({id_tarefa}) para usuário manual ({id_usuario}): {mensagem}"
    )
    response = client_singleton().put(
        f"/v1/tasks/{id_tarefa}/assignee",
        json={"task": {"assignee_id": str(id_usuario), "note": mensagem}},
    )
    if response.status_code not in (200, 204):
        bot.logger.informar(
            f"[Holmes] Falha ao alocar tarefa({id_tarefa}): status {response.status_code}"
        )
    
