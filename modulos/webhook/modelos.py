# std
import dataclasses, datetime, typing

@dataclasses.dataclass
class Document:
    name: str
    status: str
    document_id: str | None
    """Identificador do documento
    - Pode ser `None` caso não tenha sido feito o upload"""
    process_document_id: str

@dataclasses.dataclass
class Author:
    id: str
    """Identificador do autor do processo"""
    name: str
    """Nome do autor do processo"""
    email: str
    """Email do autor do processo"""

@dataclasses.dataclass
class Webhook:
    id_webhook: int
    """Identificador no webhook"""
    id_processo: str
    """Identificador do processo no Holmes"""
    tentativas: int
    """Número de tentativas no webhook"""
    controle: list[typing.Any]
    """Campo de controle do webhook"""
    criado_em: datetime.date
    """Data de criação no webhook"""
    atualizado_em: datetime.date
    """Data de atualização no webhook"""

@dataclasses.dataclass
class DadosItemWebhook:
    author: Author
    documents: list[Document]
    properties: dict[str, typing.Any]

@dataclasses.dataclass
class ItemWebhook:
    id_webhook: int
    id_processo: str
    criado_em: str
    atualizado_em: str
    tentativas: int
    controle: list
    dados: DadosItemWebhook

__all__ = [
    "Author",
    "Webhook",
    "Document",
]