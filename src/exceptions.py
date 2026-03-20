import dataclasses


class ErroNegocio(Exception):
    pass


class ErroTecnico(Exception):
    pass


@dataclasses.dataclass
class CamposNBS:
    """Campos obrigatórios para lançamento no NBS.
    Tipado e nomeado para evitar desalinhamento posicional.
    """
    cpf: str
    cod_empresa: str
    cod_filial: str
    processo: str
    iniciar_em: str
    despesa_pagamento: str
    protocolo: str
    centro_custo: str
    data_vencimento: str
    tipo_pagamento: str
