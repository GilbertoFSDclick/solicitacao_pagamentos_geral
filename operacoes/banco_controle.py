"""
Fluxo Prévio II - Inclusão de dados em Banco de Dados de Controle.
Conforme Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral.
Dados a inserir: Empresa_holmes, Filial_holmes, Cpf_holmes, Num_nf_holmes,
Serie_nf_holmes, Emissão_nf_holmes, Entrada_nf_holmes, Vlr_nf_holmes,
Protocolo_holmes, Centro_custo_uab_holmes, Dt_Vencimento_Holmes, Codigo_UAB_Contabil.
"""
import bot
from operacoes.tratar_tarefa import DadosExtraidosHolmes


def registrar_processo_controle(
    id_processo: str,
    dados_extraidos: DadosExtraidosHolmes,
    codigo_uab_contabil: str,
) -> None:
    """
    Inserir dados resgatados no Holmes + Codigo_UAB_Contabil no banco de controle.
    Doc: 'Acessar banco de dados de controle da automação. Inserir dados resgatados no Holmes.'
    Configure [banco_controle] no .ini para habilitar o insert real.
    """
    registro = {
        "id_processo": id_processo,
        "Empresa_holmes": dados_extraidos.Empresa_holmes,
        "Filial_holmes": dados_extraidos.Filial_holmes,
        "Cpf_holmes": dados_extraidos.Cpf_holmes,
        "Num_nf_holmes": dados_extraidos.Num_nf_holmes,
        "Serie_nf_holmes": dados_extraidos.Serie_nf_holmes,
        "Emissão_nf_holmes": dados_extraidos.Emissão_nf_holmes,
        "Entrada_nf_holmes": dados_extraidos.Entrada_nf_holmes,
        "Vlr_nf_holmes": dados_extraidos.Vlr_nf_holmes,
        "Protocolo_holmes": dados_extraidos.Protocolo_holmes,
        "Centro_custo_uab_holmes": dados_extraidos.Centro_custo_uab_holmes,
        "Dt_Vencimento_Holmes": dados_extraidos.Dt_Vencimento_Holmes,
        "Codigo_UAB_Contabil": codigo_uab_contabil,
    }

    habilitado = str(bot.configfile.obter_opcao_ou("banco_controle", "habilitado") or "").lower() in ("1", "true", "sim", "s")
    if habilitado:
        try:
            _inserir_registro(registro)
        except Exception as e:
            bot.logger.erro(f"[Banco Controle] Erro ao inserir: {e}")
    else:
        bot.logger.informar(
            f"[Banco Controle] Registro preparado (habilitado=false): "
            f"id_processo={id_processo}, protocolo={dados_extraidos.Protocolo_holmes}, "
            f"codigo_uab={codigo_uab_contabil}"
        )


def _inserir_registro(registro: dict) -> None:
    """
    Executa o insert no banco. Configure [banco_controle] url e tabela no .ini.
    Exemplo SQLite: url = file:controle_rpa.db
    Exemplo SQL Server: url = mssql+pyodbc://user:pass@server/db
    """
    url = bot.configfile.obter_opcao_ou("banco_controle", "url")
    tabela = bot.configfile.obter_opcao_ou("banco_controle", "tabela") or "rpa_controle_solicitacao_pagamento"
    if not url:
        bot.logger.informar("[Banco Controle] url não configurada - registro apenas logado")
        return

    try:
        import sqlalchemy
        from sqlalchemy import text
        engine = sqlalchemy.create_engine(url)
        cols = ", ".join(registro.keys())
        placeholders = ", ".join(f":{k}" for k in registro.keys())
        sql = f"INSERT INTO {tabela} ({cols}) VALUES ({placeholders})"
        with engine.connect() as conn:
            conn.execute(text(sql), registro)
            conn.commit()
        bot.logger.informar(f"[Banco Controle] Registro inserido: id_processo={registro['id_processo']}")
    except ImportError:
        bot.logger.erro("[Banco Controle] sqlalchemy não instalado. pip install sqlalchemy pyodbc (ou driver do BD)")
        raise
    except Exception as e:
        bot.logger.erro(f"[Banco Controle] Falha no insert: {e}")
        raise
