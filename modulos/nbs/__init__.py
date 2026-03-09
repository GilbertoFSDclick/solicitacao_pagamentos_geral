"""Módulo de controles de interação e manipulação do NBS"""

from .__setup import *
from .solicitacao_pagamento import (
    processar_entrada_solicitacao_pagamento,
    obter_codigo_uab_contabil,
)