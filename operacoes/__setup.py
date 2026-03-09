# std
from typing import NamedTuple
from time import sleep
from enum import Enum
from typing import Literal

# interno
import bot
import modulos
from bot.estruturas import Coordenada

# externo
class StatusOperacao(NamedTuple):
    """Classe com padrão de retorno
    - `SUCESSO` indica se a ação foi bem-sucedida
    - `MENSAGEM` passa detalhes em caso de mal-sucedida"""
    SUCESSO: bool
    MENSAGEM: str = ""

__all__ = [
    "StatusOperacao"
]