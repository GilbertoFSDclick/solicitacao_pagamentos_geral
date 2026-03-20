"""
Tratamento de tarefa para Lançamento de Solicitações de Pagamento Geral.
Mapeamento dos campos reais do Holmes (Solicitação de Pagamentos Geral).
Doc: Extrair e guardar os dados: Empresa_holmes, Filial_holmes, Cpf_holmes,
Num_nf_holmes, Serie_nf_holmes, Emissão_nf_holmes, Entrada_nf_holmes,
Vlr_nf_holmes, Protocolo_holmes, Centro_custo_uab_holmes, Dt_Vencimento_Holmes.
"""
import re
from dataclasses import dataclass
import bot
import modulos
from src.exceptions import ErroNegocio, ErroTecnico, CamposNBS
from datetime import datetime


@dataclass
class DadosExtraidosHolmes:
    """Dados a extrair da nota/processo Holmes conforme documentação atualizada."""
    Empresa_holmes: str
    Filial_holmes: str
    Cpf_holmes: str
    Num_nf_holmes: str
    Serie_nf_holmes: str
    Emissão_nf_holmes: str
    Entrada_nf_holmes: str
    Vlr_nf_holmes: str
    Protocolo_holmes: str
    Centro_custo_uab_holmes: str
    Dt_Vencimento_Holmes: str


NOME_ACAO_SUCESSO, NOME_ACAO_ERRO = bot.configfile.obter_opcoes(
    "holmes", ["nome_acao_tarefa_sucesso", "nome_acao_tarefa_erro"]
)

# Mapeamento: nome no Holmes (parcial, case-insensitive) → chave interna
# Conforme Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral
MAPEAMENTO_CAMPOS = {
    "cpf": "cpf",
    "protocolo": "protocolo",
    "processo": "processo",
    "tipo de pagamento": "tipo_pagamento",
    "valores": "despesa_pagamento",
    "devolução de adiantamento": "despesa_pagamento",
    "obrigação": "despesa_pagamento_alt",
    "observação": "observacao",
    "filial": "filial",
    "cnpj da loja": "cnpj_loja",
    "torre": "torre",
    "ait": "ait",
    "data de vencimento": "data_vencimento",
    "iniciar em": "iniciar_em",
    "centro de custo": "centro_custo",
    "centro de custo - uab - nbs": "centro_custo",
    "código empresa nbs": "codigo_empresa_nbs",
    "codigo empresa nbs": "codigo_empresa_nbs",
    "código filial nbs": "codigo_filial_nbs",
    "codigo filial nbs": "codigo_filial_nbs",
}


def _extrair_valor(propriedades: list[dict], nome_busca: str) -> str | None:
    """Extrai valor de propriedade por nome (case-insensitive, parcial)."""
    for prop in propriedades:
        nome = prop.get("name", "")
        if nome_busca.lower() in nome.lower():
            valor = prop.get("value")
            if valor is None:
                return None
            if isinstance(valor, list):
                for item in valor:
                    if isinstance(item, dict) and "text" in item:
                        return str(item["text"]).strip()
                return str(valor[0]).strip() if valor else None
            return str(valor).strip()
    return None


def _extrair_nested_centro_custo(propriedades: list[dict]) -> str | None:
    """Extrai Centro de Custo de property_values aninhado."""
    for prop in propriedades:
        if "centro de custo" in prop.get("name", "").lower():
            for pv in prop.get("property_values", []):
                if "cód. cc" in pv.get("name", "").lower() or "cod. cc" in pv.get("name", "").lower():
                    v = pv.get("value", "")
                    m = re.search(r"\d+", str(v))
                    if m:
                        return m.group()
    return None


def _obter_codigos_empresa_filial(cnpj_loja: str | None, filial_nome: str | None, torre: str) -> tuple[str | None, str | None]:
    """
    Obtém Código Empresa e Filial NBS via planilha Empresas_NBS (de/para).
    Usa CNPJ da Loja ou nome da Filial.
    """
    if not cnpj_loja and not filial_nome:
        return None, None
    try:
        planilha = bot.configfile.obter_opcao_ou("nbs", "planilha_empresas")
        if not planilha or not bot.windows.afirmar_arquivo(planilha):
            return None, None
        import pandas as pd
        df = pd.read_excel(planilha, sheet_name=torre if len(torre) <= 10 else "UAB", dtype={"CNPJ": str})
        df.columns = df.columns.str.strip()
        cnpj_limpo = re.sub(r"\D", "", str(cnpj_loja or "")) if cnpj_loja else ""
        if cnpj_limpo:
            mask = df["CNPJ"].astype(str).str.replace(r"\D", "", regex=True) == cnpj_limpo.zfill(14)
            if mask.any():
                row = df[mask].iloc[0]
                cod_emp = str(row.get("Nbs_empresa", row.get("Cod_empresa", ""))).strip()
                cod_fil = str(row.get("Nbs_Filial", "")).strip()
                return cod_emp or None, cod_fil or None
        if filial_nome:
            cod_fil = filial_nome.split("-")[0].strip() if "-" in str(filial_nome) else filial_nome[:3]
            return None, cod_fil
    except Exception as e:
        bot.logger.erro(f"Erro ao obter códigos Empresa/Filial: {e}")
    return None, None


def _normalizar_data_holmes(valor: str | None) -> str | None:
    """Normaliza data para DD/MM/YYYY removendo qualquer parte de hora."""
    if valor is None:
        return None
    s = str(valor).strip()
    if not s:
        return None

    # ISO (ex.: 2026-03-17T09:28:13.000Z)
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00")).strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        pass

    formatos = (
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    )
    for fmt in formatos:
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            continue

    m_br = re.search(r"(\d{2}/\d{2}/\d{4})", s)
    if m_br:
        return m_br.group(1)

    m_iso = re.search(r"(\d{4}-\d{2}-\d{2})", s)
    if m_iso:
        try:
            return datetime.strptime(m_iso.group(1), "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            pass

    dig = re.sub(r"\D", "", s)
    if len(dig) >= 8:
        base = dig[:8]
        if base.startswith("20"):
            try:
                return datetime.strptime(base, "%Y%m%d").strftime("%d/%m/%Y")
            except ValueError:
                pass
        try:
            return datetime.strptime(base, "%d%m%Y").strftime("%d/%m/%Y")
        except ValueError:
            pass

    return None


def tratar_tarefa_aberta(identificador: str) -> tuple[CamposNBS, dict]:
    """
    Obtém informações da tarefa e processo do Holmes.
    Campos Holmes reais: Protocolo, Filial, CPF, Tipo de pagamento, Valores,
    Devolução de adiantamento/Obrigação, Observação, CNPJ da Loja, Torre, AIT.
    Retorna (campos_obrigatorios, dados_processo).
    """
    bot.logger.informar(f"Iniciando tratamento da tarefa ({identificador})")

    campos_invalidos: list[str] = []

    try:
        tarefa = modulos.holmes.consulta_tarefa(identificador)
        process_id = tarefa.get("process_id")
        processo = modulos.holmes.consulta_processo(process_id)

        propriedades_tarefa: list[dict] = tarefa.get("properties", [])
        propriedades_processo: list[dict] = processo["instance"]["property_values"]
        acoes: list[dict] = tarefa.get("actions", [])

        def _get(nome: str) -> str | None:
            v = _extrair_valor(propriedades_processo, nome)
            return v or _extrair_valor(propriedades_tarefa, nome)

        # Validar ações
        # Em alguns fluxos/etapas o Holmes expõe "Aprovar" ou "Avançar"
        # como ação de sucesso, mesmo quando o nome configurado é diferente.
        acaoSucessoID = acaoErroID = None
        for acao in acoes:
            nome, _id = str(acao.get("name") or ""), acao.get("id")
            nome_norm = bot.util.normalizar(nome)
            sucesso_norm = bot.util.normalizar(str(NOME_ACAO_SUCESSO or ""))
            erro_norm = bot.util.normalizar(str(NOME_ACAO_ERRO or ""))

            if nome_norm in {sucesso_norm, "aprovar", "avancar"}:
                acaoSucessoID = _id
            elif nome_norm in {erro_norm, "pendenciar"}:
                acaoErroID = _id
        assert None not in (acaoSucessoID, acaoErroID), (
            f"A tarefa ({identificador}) não possui ações esperadas: "
            f"[{NOME_ACAO_SUCESSO}, {NOME_ACAO_ERRO}]"
        )

        # Campos conforme Holmes real (CPF ou CNPJ para Pessoa Jurídica)
        cpf = _get("cpf") or _get("cnpj")
        if cpf:
            cpf = re.sub(r"\D", "", cpf)
        if not cpf:
            campos_invalidos.append("CPF")

        protocolo = _get("protocolo")
        if not protocolo:
            campos_invalidos.append("Protocolo")

        tipo_pagamento = _get("tipo de pagamento")
        if not tipo_pagamento:
            campos_invalidos.append("Tipo de pagamento")

        despesa_pagamento = _get("valores") or _get("devolução de adiantamento") or _get("obrigação")
        if despesa_pagamento:
            despesa_pagamento = despesa_pagamento.replace(",", ".")
        if not despesa_pagamento:
            campos_invalidos.append("Valores / Despesa Pagamento")

        observacao = _get("observação") or protocolo or ""

        filial_nome = _get("filial")
        cnpj_loja = _get("cnpj da loja")
        if cnpj_loja:
            cnpj_loja = re.sub(r"\D", "", cnpj_loja)
        torre = _get("torre") or "UAB"
        if len(torre) > 20:
            torre = "UAB"

        # Doc: Código Empresa/Filial NBS do Holmes têm prioridade
        cod_empresa = _get("código empresa nbs") or _get("codigo empresa nbs")
        cod_filial = _get("código filial nbs") or _get("codigo filial nbs")
        if not cod_empresa or not cod_filial:
            cod_emp, cod_fil = _obter_codigos_empresa_filial(cnpj_loja, filial_nome, torre)
            cod_empresa = cod_empresa or cod_emp
            cod_filial = cod_filial or cod_fil
        if not cod_empresa and not cod_filial:
            if filial_nome and "-" in str(filial_nome):
                cod_filial = filial_nome.split("-")[0].strip()
        if not cod_empresa:
            campos_invalidos.append("Código Empresa NBS")
        if not cod_filial:
            campos_invalidos.append("Código Filial NBS")

        # Doc: Número NF = campo Processo Holmes
        processo_num = _get("processo") or _get("ait") or protocolo or ""
        if processo_num:
            processo_num = processo_num.replace("-", "").replace(".", "").replace(" ", "")

        iniciar_em = _normalizar_data_holmes(_get("iniciar em"))
        if not iniciar_em:
            iniciar_em = datetime.now().strftime("%d/%m/%Y")

        centro_custo = _get("centro de custo") or _extrair_nested_centro_custo(propriedades_processo)
        if not centro_custo:
            centro_custo = "0"

        data_vencimento_raw = _get("data de vencimento") or _get("data de pagamento")
        data_vencimento = _normalizar_data_holmes(data_vencimento_raw)
        if not data_vencimento:
            campos_invalidos.append("Data de Vencimento")

        assert not campos_invalidos, f"Campos obrigatórios ausentes: {campos_invalidos}"

        campos = CamposNBS(
            cpf=cpf,
            cod_empresa=cod_empresa,
            cod_filial=cod_filial,
            processo=processo_num,
            iniciar_em=iniciar_em,
            despesa_pagamento=despesa_pagamento,
            protocolo=protocolo,
            centro_custo=centro_custo,
            data_vencimento=data_vencimento,
            tipo_pagamento=tipo_pagamento,
        )

        dados_extraidos = DadosExtraidosHolmes(
            Empresa_holmes=cod_empresa,
            Filial_holmes=cod_filial,
            Cpf_holmes=cpf,
            Num_nf_holmes=processo_num,
            Serie_nf_holmes="E",
            Emissão_nf_holmes=iniciar_em,
            Entrada_nf_holmes=iniciar_em,
            Vlr_nf_holmes=despesa_pagamento,
            Protocolo_holmes=protocolo,
            Centro_custo_uab_holmes=centro_custo,
            Dt_Vencimento_Holmes=data_vencimento,
        )

        bot.logger.informar(
            f"Dados extraídos: protocolo={protocolo}, tipo={tipo_pagamento}, valor={despesa_pagamento}"
        )

        return (campos, {"propriedades_processo": propriedades_processo, "dados_extraidos": dados_extraidos})

    except AssertionError as e:
        raise ErroNegocio(str(e))
    except Exception as e:
        raise ErroTecnico(f"Erro ao obter dados do Holmes: {e}")
