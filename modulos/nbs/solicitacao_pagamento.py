"""
Fluxo NBS para Lançamento de Solicitações de Pagamento Geral.
Conforme Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral.
Admin → NBS Fiscal → Entrada (Só Diversa).
"""
import bot
import pandas as pd
from time import sleep
from datetime import datetime
from modulos.interface import Elemento

from .__setup import ErroNBS, RetornoStatus

# Imagens para fallback quando pywinauto falhar (ver GUIA_CAPTURA_IMAGENS_NBS.md)
IMAGENS_FALLBACK = {
    "aba_contabilizacao": "imagens/nbs_aba_contabilizacao.PNG",
    "botao_incluir": "imagens/nbs_botao_incluir.PNG",
    "incluir_entrada_diversa": "imagens/nbs_incluir_entrada_diversa.PNG",
    "aba_faturamento": "imagens/nbs_aba_faturamento.PNG",
    "botao_gerar": "imagens/nbs_botao_gerar.PNG",
    "botao_confirmar": "imagens/nbs_botao_confirmar.PNG",
    "cnpj_nao_encontrado": "imagens/nbs_cnpj_nao_encontrado.PNG",
}


def _encontrar_imagem(caminho: str, confianca: float = 0.8, segundos: int = 5):
    """Retorna coordenada da imagem se encontrada, None caso contrário."""
    try:
        return bot.imagem.procurar_imagem(caminho, confianca=confianca, segundos=segundos)
    except Exception:
        return None


def _clicar_por_imagem_se_existir(caminho: str, confianca: float = 0.8, segundos: int = 5) -> bool:
    """Fallback: clica por reconhecimento de imagem se o arquivo existir e for encontrado."""
    img = _encontrar_imagem(caminho, confianca, segundos)
    if img:
        bot.mouse.clicar_mouse(coordenada=img)
        bot.logger.informar(f"Fallback imagem: {caminho}")
        return True
    return False


def obter_codigo_uab_contabil(tipo_pagamento_holmes: str) -> str | None:
    """
    TE07: Consulta planilha TIPOS_DE_PAGAMENTO.
    Coluna A = Tipo de Pagamento Holmes → Coluna D = Código UAB (CODIGO_UAB_CONTABIL).
    Retorna None se não encontrado (TE07: não lançar, Business Exception).
    """
    try:
        caminho = bot.configfile.obter_opcao_ou("nbs", "planilha_tipos_pagamento")
        if not caminho or not bot.windows.afirmar_arquivo(caminho):
            bot.logger.erro("Planilha TIPOS_DE_PAGAMENTO não encontrada ou inacessível")
            return None

        # Doc: TIPOS_DE_PAGAMENTO em DIRETORIO_PLANILHA, aba Planilha1 (ou Plan1)
        # Coluna A = Tipos de Pagamentos, Coluna D = UAB (CODIGO_UAB_CONTABIL)
        try:
            df = pd.read_excel(caminho, sheet_name="Plan1", header=1)
        except Exception:
            try:
                df = pd.read_excel(caminho, sheet_name="Planilha1", header=1)
            except Exception:
                df = pd.read_excel(caminho, sheet_name=0, header=1)
        # Coluna A (0) = Tipos de Pagamentos, Coluna D (3) = UAB
        if df.shape[1] < 4:
            bot.logger.erro("Planilha TIPOS_DE_PAGAMENTO sem coluna D")
            return None

        col_a = df.iloc[:, 0].astype(str).str.strip().str.lower()
        col_d = df.iloc[:, 3].astype(str)
        tipo_busca = str(tipo_pagamento_holmes).strip().lower()

        mask = col_a == tipo_busca
        if mask.any():
            codigo = str(col_d[mask].iloc[0]).strip()
            if codigo and codigo.lower() != "nan":
                bot.logger.informar(f"Tipo Pagamento '{tipo_pagamento_holmes}' → Código UAB: {codigo}")
                return codigo

        bot.logger.informar(f"TE07: Tipo de Pagamento '{tipo_pagamento_holmes}' não encontrado na planilha")
        return None
    except Exception as e:
        bot.logger.erro(f"Erro ao ler planilha TIPOS_DE_PAGAMENTO: {e}")
        return None


def processar_entrada_solicitacao_pagamento(
    campos_obrigatorios: list,
    dados_processo: dict,
    parar_antes_confirmar: bool = False,
    parar_apos_observacoes: bool = False,
    codigo_uab_contabil: str | None = None,
) -> RetornoStatus:
    """
    Fluxo NBS: Admin → NBS Fiscal → Entrada Só Diversa.
    campos_obrigatorios: [cpf, cod_empresa, cod_filial, processo, iniciar_em, despesa_pagamento,
                         protocolo, centro_custo, data_vencimento, tipo_pagamento]
    parar_antes_confirmar: Se True, para antes de clicar em Confirmar (nao cria entrada).
    parar_apos_observacoes: Se True, para após Total Nota e Observações (para testes parciais).
    """
    (
        cpf,
        cod_empresa,
        cod_filial,
        processo,
        iniciar_em,
        despesa_pagamento,
        protocolo,
        centro_custo,
        data_vencimento,
        tipo_pagamento,
    ) = campos_obrigatorios

    # TE07: Tipo de Pagamento não encontrado na planilha
    codigo_uab = codigo_uab_contabil or obter_codigo_uab_contabil(tipo_pagamento)
    if not codigo_uab:
        return RetornoStatus(
            False,
            f"Tipo de Pagamento '{tipo_pagamento}' não parametrizado na planilha TIPOS_DE_PAGAMENTO. "
            "Não lançar no NBS. Registrar Business Exception.",
        )

    # Data de Corte: validar antes de abrir o NBS (evita trabalho desnecessário)
    data_corte = bot.configfile.obter_opcao_ou("nbs", "data_corte")
    bloquear_fora_corte = str(bot.configfile.obter_opcao_ou("nbs", "data_corte_bloquear") or "").lower() in ("1", "true", "sim", "s")
    if data_corte and bloquear_fora_corte:
        try:
            dt_corte = datetime.strptime(str(data_corte).strip(), "%d/%m/%Y")
            dt_venc = datetime.strptime(str(data_vencimento).strip(), "%d/%m/%Y")
            if dt_venc > dt_corte:
                msg = f"Data de Vencimento ({data_vencimento}) fora da Data de Corte ({data_corte})."
                bot.logger.alertar(msg)
                return RetornoStatus(False, msg)
        except (ValueError, TypeError):
            pass

    # Formatar data Iniciar em para Emissão/Entrada (YYYY-MM-DD ou DD/MM/YYYY)
    try:
        dt = datetime.fromisoformat(iniciar_em.replace("Z", "+00:00"))
        data_emissao = dt.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        data_emissao = str(iniciar_em)

    try:
        bot.logger.informar(f"Iniciando entrada Solicitação Pagamento - processo {processo}, protocolo {protocolo}")
        bot.estruturas.Janela("NBS Sistema Financeiro").focar()

        # Aguardar janela inicial (NBS Sistema Financeiro ou similar)
        bot.util.aguardar_condicao(
            lambda: "Sistema Financeiro" in bot.windows.Janela.titulos_janelas() or "SisFin" in str(bot.windows.Janela.titulos_janelas()), 15
        )
        titulos = bot.windows.Janela.titulos_janelas()
        janela = bot.windows.Janela("NBS Sistema Financeiro") if "NBS" in str(titulos) else bot.windows.Janela()
        if "SisFin" not in str(titulos):
            elemento = janela.elementos(class_name="TOvcPictureField", top_level_only=False)
            if elemento:
                bot.mouse.clicar_mouse(coordenada=Elemento(elemento[0]).coordenada)
                bot.teclado.digitar_teclado("SisFin")
                bot.teclado.apertar_tecla("enter")
                sleep(2)

        # Aguardar Empresa/Filial/Exercício Contábil OU Sistema Financeiro - SISFIN
        def _janela_empresa_filial():
            t = bot.windows.Janela.titulos_janelas()
            return "Empresa/Filial" in str(t) or "Exercício Contábil" in str(t) or "Sistema Financeiro - SISFIN" in str(t)

        bot.util.aguardar_condicao(_janela_empresa_filial, 15)
        titulos = bot.windows.Janela.titulos_janelas()
        if "Empresa" in str(titulos) or "Exercício" in str(titulos):
            janela_ef = None
            for t in titulos:
                if "Empresa" in t or "Exercício" in t:
                    janela_ef = bot.windows.Janela(t)
                    break
            if not janela_ef:
                janela_ef = bot.windows.Janela()
            elementos = janela_ef.elementos(class_name="TOvcPictureField", top_level_only=False)
            if not elementos:
                elementos = janela_ef.elementos(class_name="TEdit", top_level_only=False)
            if len(elementos) >= 2:
                bot.mouse.clicar_mouse(coordenada=Elemento(elementos[0]).coordenada)
                bot.teclado.digitar_teclado(cod_empresa)
                bot.teclado.apertar_tecla("tab")
                bot.teclado.digitar_teclado(cod_filial)
                bot.teclado.apertar_tecla("tab")
            else:
                bot.teclado.digitar_teclado(cod_empresa)
                bot.teclado.apertar_tecla("tab")
                bot.teclado.digitar_teclado(cod_filial)
            # Clicar Confirma (verde) ou Enter
            for btn in janela_ef.elementos(class_name="TBitBtn", top_level_only=False):
                if "confirma" in str(btn.window_text()).lower():
                    bot.mouse.clicar_mouse(coordenada=Elemento(btn).coordenada)
                    break
            else:
                bot.teclado.apertar_tecla("enter")
        else:
            janela = bot.windows.Janela("Sistema Financeiro - SISFIN")
            elementos = janela.elementos(class_name="TOvcPictureField", top_level_only=False)
            if len(elementos) >= 2:
                bot.mouse.clicar_mouse(coordenada=Elemento(elementos[0]).coordenada)
                bot.teclado.digitar_teclado(cod_empresa)
                bot.teclado.apertar_tecla("tab")
                bot.teclado.digitar_teclado(cod_filial)
                bot.teclado.apertar_tecla("tab")
                bot.teclado.apertar_tecla("enter")
            else:
                bot.teclado.digitar_teclado(cod_empresa)
                bot.teclado.apertar_tecla("tab")
                bot.teclado.digitar_teclado(cod_filial)
                bot.teclado.apertar_tecla("enter")
        sleep(2)

        # Sistema Fiscal: Admin → NBS Fiscal (ou já está em Sistema Fiscal)
        bot.util.aguardar_condicao(
            lambda: "Sistema Fiscal" in str(bot.windows.Janela.titulos_janelas()) or "Entradas" in str(bot.windows.Janela.titulos_janelas()) or "Entrada" in str(bot.windows.Janela.titulos_janelas()), 10
        )
        titulos = bot.windows.Janela.titulos_janelas()
        if "Entradas" not in str(titulos) and "Entrada Diversa" not in str(titulos):
            bot.logger.informar("Acessando Admin → NBS Fiscal")
            janela_fiscal = bot.windows.Janela("Sistema Fiscal") if "Sistema Fiscal" in str(titulos) else bot.windows.Janela()
            try:
                itens = janela_fiscal.elementos(class_name="TMenuItem", top_level_only=False)
                for item in itens:
                    if "fiscal" in str(item.window_text()).lower() or "nbs" in str(item.window_text()).lower():
                        bot.mouse.clicar_mouse(coordenada=Elemento(item).coordenada)
                        break
            except Exception:
                pass
            bot.teclado.digitar_teclado("n")
            bot.teclado.digitar_teclado("b")
            bot.teclado.digitar_teclado("s")
            sleep(1)
            bot.teclado.apertar_tecla("enter")
            sleep(2)

        # Incluir Entrada (Só Diversa) - segundo ícone na toolbar
        bot.logger.informar("Clicando em Incluir Entrada (Só Diversa)")
        clicou_incluir = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["incluir_entrada_diversa"], segundos=5)
        if not clicou_incluir:
            bot.teclado.apertar_tecla("alt")
            bot.teclado.digitar_teclado("e")
            sleep(1)
            bot.teclado.apertar_tecla("down")
            bot.teclado.apertar_tecla("down")
            bot.teclado.apertar_tecla("enter")
        sleep(3)

        # Aguardar janela de entrada
        titulos = bot.windows.Janela.titulos_janelas()
        janela_entrada = None
        for t in titulos:
            if "Entrada" in t or "Diversa" in t:
                janela_entrada = bot.windows.Janela(t)
                break
        if not janela_entrada:
            janela_entrada = bot.windows.Janela()

        # Fornecedor CNPJ/CPF (TE01: se não encontrado, pendenciar)
        bot.logger.informar("Preenchendo Fornecedor CPF")
        campos = janela_entrada.elementos(class_name="TEdit", top_level_only=False)
        if campos:
            bot.mouse.clicar_mouse(coordenada=Elemento(campos[0]).coordenada)
        bot.teclado.digitar_teclado(cpf)
        bot.teclado.apertar_tecla("tab")  # Trazer registro (doc: Dar Tab para trazer o registro)
        sleep(2)

        # TE01: CNPJ/CPF não encontrado → Pendenciar tarefa Holmes. Motivo: CNPJ não encontrado.
        titulos = str(bot.windows.Janela.titulos_janelas()).lower()
        img_erro = _encontrar_imagem(IMAGENS_FALLBACK["cnpj_nao_encontrado"], segundos=2)
        if "não encontrado" in titulos or "não cadastrado" in titulos or "nao encontrado" in titulos or "nao cadastrado" in titulos or img_erro:
            if img_erro:
                bot.teclado.apertar_tecla("enter")
            raise ErroNBS("CNPJ não encontrado.")

        # Número NF, Série E, Emissão, Entrada (doc: Processo→Nº NF, Série=E, Iniciar em→Emissão e Entrada)
        bot.logger.informar("Preenchendo Número NF, Série, Emissão, Entrada")
        for _ in range(2):
            bot.teclado.apertar_tecla("tab")
        bot.teclado.digitar_teclado(processo)
        bot.teclado.apertar_tecla("tab")
        bot.teclado.digitar_teclado("E")  # Série E
        bot.teclado.apertar_tecla("tab")
        bot.teclado.digitar_teclado(data_emissao)
        bot.teclado.apertar_tecla("tab")
        bot.teclado.digitar_teclado(data_emissao)  # Entrada = Iniciar em
        for _ in range(4):
            bot.teclado.apertar_tecla("tab")

        # Desmarcar "Quero esta nota no livro fiscal" (doc: Desmarcar Flag)
        bot.teclado.apertar_tecla("space")
        sleep(0.5)

        # Total Nota (Despesa Pagamento) e Observações (Protocolo)
        bot.teclado.digitar_teclado(despesa_pagamento)
        bot.teclado.apertar_tecla("tab")
        bot.teclado.digitar_teclado(protocolo)
        sleep(1)

        if parar_apos_observacoes:
            bot.logger.informar("PARADO: Preenchido até Total Nota e Observações. Continuar depois.")
            return RetornoStatus(True, "Parado após Observações - aguardando continuação.")

        # Guia Contabilização → + (Incluir)
        bot.logger.informar("Preenchendo Contabilização")
        clicou_aba = False
        abas = janela_entrada.elementos(class_name="TTabSheet", top_level_only=False)
        for aba in abas:
            if "contabil" in str(aba.window_text()).lower():
                bot.mouse.clicar_mouse(coordenada=Elemento(aba).coordenada)
                clicou_aba = True
                break
        if not clicou_aba:
            clicou_aba = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["aba_contabilizacao"])
        if not clicou_aba:
            raise ErroNBS("Não foi possível localizar aba Contabilização")
        sleep(1)
        clicou_incluir = False
        botoes = janela_entrada.elementos(class_name="TBitBtn", top_level_only=False)
        for btn in botoes:
            if "+" in str(btn.window_text()) or "incluir" in str(btn.window_text()).lower():
                bot.mouse.clicar_mouse(coordenada=Elemento(btn).coordenada)
                clicou_incluir = True
                break
        if not clicou_incluir:
            clicou_incluir = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["botao_incluir"])
        if not clicou_incluir:
            raise ErroNBS("Não foi possível localizar botão Incluir contabilização")
        sleep(2)

        # Preencher Conta Contábil (código UAB) e Centro de Custo
        bot.util.aguardar_condicao(
            lambda: "Incluir" in str(bot.windows.Janela.titulos_janelas()) or "Conta" in str(bot.windows.Janela.titulos_janelas()),
            5
        )
        janela_contab = bot.windows.Janela()
        for t in bot.windows.Janela.titulos_janelas():
            if "Conta" in t or "Contabilização" in t:
                janela_contab = bot.windows.Janela(t)
                break
        campos_contab = janela_contab.elementos(class_name="TEdit", top_level_only=False)
        if len(campos_contab) >= 1:
            bot.mouse.clicar_mouse(coordenada=Elemento(campos_contab[0]).coordenada)
            bot.teclado.digitar_teclado(codigo_uab)
            bot.teclado.apertar_tecla("enter")
        bot.teclado.apertar_tecla("tab")
        bot.teclado.digitar_teclado(centro_custo)
        bot.teclado.apertar_tecla("enter")
        sleep(1)
        bot.teclado.apertar_tecla("enter")
        sleep(1)

        # Confirmar Contabilização
        botoes_ok = janela_contab.elementos(class_name="TBitBtn", top_level_only=False)
        for btn in botoes_ok:
            if "confirmar" in str(btn.window_text()).lower() or "ok" in str(btn.window_text()).lower():
                bot.mouse.clicar_mouse(coordenada=Elemento(btn).coordenada)
                break
        sleep(2)

        # Guia Faturamento
        bot.logger.informar("Preenchendo Faturamento")
        clicou_fat = False
        for aba in janela_entrada.elementos(class_name="TTabSheet", top_level_only=False):
            if "faturamento" in str(aba.window_text()).lower():
                bot.mouse.clicar_mouse(coordenada=Elemento(aba).coordenada)
                clicou_fat = True
                break
        if not clicou_fat:
            clicou_fat = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["aba_faturamento"])
        if not clicou_fat:
            raise ErroNBS("Não foi possível localizar aba Faturamento")
        sleep(1)
        # Total Parcelas = 1
        numeros = janela_entrada.elementos(class_name="TOvcNumericField", top_level_only=False)
        if numeros:
            bot.mouse.clicar_mouse(coordenada=Elemento(numeros[0]).coordenada)
            bot.teclado.apertar_tecla("backspace")
            bot.teclado.digitar_teclado("1")
        clicou_gerar = False
        for btn in janela_entrada.elementos(class_name="TBitBtn", top_level_only=False):
            if "gerar" in str(btn.window_text()).lower():
                bot.mouse.clicar_mouse(coordenada=Elemento(btn).coordenada)
                clicou_gerar = True
                break
        if not clicou_gerar:
            clicou_gerar = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["botao_gerar"])
        if not clicou_gerar:
            raise ErroNBS("Não foi possível localizar botão Gerar")
        sleep(2)

        # Tipo Pagamento = Boleto, Natureza Despesa = Outras Despesas
        combos = janela_entrada.elementos(class_name="TwwDBLookupCombo", top_level_only=False)
        if len(combos) >= 2:
            bot.mouse.clicar_mouse(coordenada=Elemento(combos[1]).coordenada)
            bot.teclado.digitar_teclado("Boleto")
            bot.teclado.apertar_tecla("tab")
            bot.mouse.clicar_mouse(coordenada=Elemento(combos[0]).coordenada)
            bot.teclado.digitar_teclado("Outras Despesas")
            bot.teclado.apertar_tecla("enter")
        sleep(1)

        # Vencimento ← Data de Vencimento Holmes (validação Data de Corte já feita no início)
        datas = janela_entrada.elementos(class_name="TOvcDbPictureField", top_level_only=False)
        if datas:
            bot.mouse.clicar_mouse(coordenada=Elemento(datas[0]).coordenada)
            bot.teclado.apertar_tecla("right", quantidade=10)
            bot.teclado.apertar_tecla("backspace", quantidade=10)
            bot.teclado.digitar_teclado(data_vencimento)

        if parar_antes_confirmar:
            bot.logger.informar("PARADO: Formulario preenchido. Nao clicou em Confirmar (nao cria entrada).")
            return RetornoStatus(True, "Teste: parado antes de Confirmar.")

        # Confirmar
        clicou_confirmar = False
        for btn in janela_entrada.elementos(class_name="TBitBtn", top_level_only=False):
            if "confirmar" in str(btn.window_text()).lower():
                bot.mouse.clicar_mouse(coordenada=Elemento(btn).coordenada)
                clicou_confirmar = True
                break
        if not clicou_confirmar:
            clicou_confirmar = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["botao_confirmar"])
        if not clicou_confirmar:
            raise ErroNBS("Não foi possível localizar botão Confirmar")
        sleep(3)

        # Pop-up Aviso NF-e: guardar num_controle, OK
        num_controle = None
        if "Aviso" in str(bot.windows.Janela.titulos_janelas()) or "Informação" in str(bot.windows.Janela.titulos_janelas()):
            bot.teclado.atalho_teclado(["ctrl", "c"])
            num_controle = bot.teclado.texto_copiado()
            bot.teclado.apertar_tecla("enter")

        # Ficha de Controle de Pagamento → Cancelar
        sleep(2)
        if "Ficha" in str(bot.windows.Janela.titulos_janelas()) or "Controle" in str(bot.windows.Janela.titulos_janelas()):
            for t in bot.windows.Janela.titulos_janelas():
                if "Ficha" in t or "Controle" in t:
                    janela_ficha = bot.windows.Janela(t)
                    janela_ficha.fechar()
                    break

        bot.logger.informar(f"Processamento concluído. num_controle={num_controle}")
        return RetornoStatus(True)

    except ErroNBS as e:
        return RetornoStatus(False, str(e))
    except Exception as e:
        bot.logger.alertar(f"Erro no processamento: {e}")
        return RetornoStatus(False, str(e))
