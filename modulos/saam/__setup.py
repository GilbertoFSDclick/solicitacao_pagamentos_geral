import bot
import base64

NOME_ARQUIVO_XML = "tmp_nfe.xml"

'https://www.saamauditoria-1.com.br:8444/WebServiceSAAM-API/webresources/servicoClientes/getXmlPelaChave?cnpj=11111111111111&chave=52222222222222222222222222222222222222222222'

class ErroSAAM (Exception):
    """Erro próprio para o sistema SAAM"""
    def __init__ (self, mensagem: str) -> None:
        self.mensagem = mensagem
        super().__init__(self.mensagem)

def consulta_xml(cnpj: str, chave: str) -> bool | Exception:
    """Consultar tarefa `identificador`
    - Variáveis utilizadas `[saam] -> "host", "token"`
    - Exceção `SAAM` caso ocorra algum problema"""
    bot.logger.informar(f"Consultando xml da chave({ chave })")
    # obter as variáveis de ambiente
    host, token = bot.configfile.obter_opcoes("saam", ["host", "token"])

    # realizar a chamada HTTP
    url = rf"{ host }/WebServiceSAAM-API/webresources/servicoClientes/getXmlPelaChave?cnpj={ cnpj }&chave={ chave }"
    headers = { "token": token, "Accept": "application/json" }
    response = bot.http.request("GET", url, headers=headers, timeout=120)

    # validações
    if response.status_code != 200:
        raise ErroSAAM(f"O status code '{ response.status_code }' foi diferente do esperado")
    
    body = response.json()
    if not isinstance(body, dict):
        raise ErroSAAM(f"O conteúdo json não possui o formato esperado")
    
    xml_base_64 = body.get('xml')
    if not xml_base_64: raise Exception(f"Resultado da Consulta a nota diferente do esperado: {body}")
    xml_decode = base64.b64decode(xml_base_64)

    with open(NOME_ARQUIVO_XML, 'wb') as arquivo:
        arquivo.write(xml_decode)

    # OK
    bot.logger.informar(f"Consulta realizada com sucesso")

    return True

def consulta_xml_nota(cnpj: str, num_nota: str, cnpj_emit: str, data_emissao: str) -> bool | Exception:
    """Consultar tarefa `identificador`
    - Variáveis utilizadas `[saam] -> "host", "token"`
    - Exceção `SAAM` caso ocorra algum problema"""
    bot.logger.informar(f"Consultando xml da nota({ num_nota })")
    # obter as variáveis de ambiente
    host, token = bot.configfile.obter_opcoes("saam", ["host", "token"])

    # realizar a chamada HTTP
    
    url = rf"{ host }/WebServiceSAAM-API/webresources/servicoClientes/getXmlPeloModeloCnpjPartNumDocDataDoc?cnpj={cnpj}&modelo=55&cnpjPart={cnpj_emit}&numDoc={num_nota}&dtDoc={data_emissao}"
    headers = { "token": token, "Accept": "application/json" }
    response = bot.http.request("GET", url, headers=headers, timeout=120)

    # validações
    if response.status_code != 200:
        raise ErroSAAM(f"O status code '{ response.status_code }' foi diferente do esperado")
    
    body = response.json()
    if not isinstance(body, dict):
        raise ErroSAAM(f"O conteúdo json não possui o formato esperado")
    
    xml_base_64 = body.get('xml')
    if not xml_base_64: raise Exception(f"Resultado da Consulta a nota diferente do esperado: {body}")
    xml_decode = base64.b64decode(xml_base_64)

    with open(NOME_ARQUIVO_XML, 'wb') as arquivo:
        arquivo.write(xml_decode)

    # OK
    bot.logger.informar(f"Consulta realizada com sucesso")

    return True