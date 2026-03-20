# MÓDULO SAAM — não utilizado neste processo (Lançamento Solicitação Pagamento Geral).
# Removido do agregador modulos/__init__.py. Mantido como stub para não quebrar
# referências externas eventuais.


class ErroSAAM(Exception):
    """Erro próprio para o sistema SAAM (legado)."""


def consulta_xml(*_args, **_kwargs):
    raise ErroSAAM("Módulo SAAM desativado para este processo")


def consulta_xml_nota(*_args, **_kwargs):
    raise ErroSAAM("Módulo SAAM desativado para este processo")


__all__ = ["ErroSAAM", "consulta_xml", "consulta_xml_nota"]
