# std
from datetime import datetime

# interno
import bot

class ErroDashboard (Exception):
    """Erro próprio para o Dashboard"""
    def __init__ (self, mensagem: str) -> None:
        self.mensagem = mensagem
        super().__init__(self.mensagem)


def gerar_estatistica (chave: str, protocolo: str, mensagem="") -> None:
    """Gerar estatística com as ações executadas nas tarefas para dashboard"""
    bot.logger.informar("Enviando requisição para a API do dashboard para gerar estatística")

    data_hora = datetime.today().strftime("%Y-%m-%dT%H:%M:%S%z")
    # realizar a chamada HTTP
    url = rf"https://rpa-admin.dclick.com.br/admin/api/Dashboard/GravarTransacao"
    
    # headers = { "api_token": token, "Accept": "application/json" }
    body = {
        "CodigoAutomacao": "ORI_PRODNBS",
        "Chave": chave, 
        "DataHora": data_hora,
        "Identificador": protocolo,
        "Mensagem": mensagem[:200]
    }

    response = bot.http.request("POST", url, json=body, timeout=120)

    bot.logger.debug(f"""Resposta da chamada:
        codigo: { response.status_code }
        headers: { response.headers }
        body: { response.text }""")

    # validações
    if response.status_code != 200:
        raise ErroDashboard(f"O status code '{ response.status_code }' foi diferente do esperado")
    
    body = response.json()
    if not isinstance(body, dict):
        raise ErroDashboard(f"O conteúdo json não possui o formato esperado")
    
    # OK
    bot.logger.informar(f"Estatística gerada com sucesso")

    return body

__all__ = [
    "ErroDashboard",
    "gerar_estatistica"
]