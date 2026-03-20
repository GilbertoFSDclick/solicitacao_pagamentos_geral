"""
Fluxo NBS para Lançamento de Solicitações de Pagamento Geral.
Conforme Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral.
Admin → NBS Fiscal → Entrada (Só Diversa).
"""
import bot
import pandas as pd
import re
from time import sleep
from datetime import datetime
from pathlib import Path
from typing import Callable
from modulos.interface import Elemento
from src.exceptions import CamposNBS

from .__setup import ErroNBS, RetornoStatus

# Coordenada do botão Nota Fiscal Entradas (captura em VM via MouseInfo).
# Pode ser sobrescrita por config em [nbs]:
# - nota_fiscal_entradas_xy = 372,213
# ou
# - nota_fiscal_entradas_x = 372
# - nota_fiscal_entradas_y = 213
NOTA_FISCAL_ENTRADAS_XY = (372, 213)

# Coordenada do botão Incluir Entrada (Só Diversa) na tela Entradas.
# Pode ser sobrescrita por config em [nbs]:
# - incluir_entrada_diversa_xy = 211,81
# ou
# - incluir_entrada_diversa_x = 211
# - incluir_entrada_diversa_y = 81
INCLUIR_ENTRADA_DIVERSA_XY = (377,213)

# Imagens para fallback quando pywinauto falhar (ver GUIA_CAPTURA_IMAGENS_NBS.md)
IMAGENS_FALLBACK = {
    "aba_contabilizacao": "imagens/contabilizacao.PNG",
    "botao_incluir": "imagens/incluir_contabilizacao.PNG",
    "incluir_entrada_diversa": "imagens/nbs_incluir_entrada_diversa.png",
    "nota_fiscal_entradas": "imagens/nbs_nota_fiscal_entradas.png",
    "aba_faturamento": "imagens/faturamento.PNG",
    "botao_gerar": "imagens/gerar.PNG",
    "botao_confirmar": "imagens/confirmar.PNG",
    "cnpj_nao_encontrado": "imagens/fornecedor_nao_cadastrado.PNG",
}


def _diagnostico_ui_ativo() -> bool:
    """Habilita logs/snapshots extras via .ini ([nbs] diagnostico_ui=true)."""
    try:
        valor = bot.configfile.obter_opcao_ou("nbs", "diagnostico_ui")
        return str(valor or "").strip().lower() in ("1", "true", "sim", "s", "yes", "on")
    except Exception:
        return False


def _diagnostico_ui_snapshot(contexto: str, capturar_print: bool = True) -> None:
    """Registra títulos de janelas e, opcionalmente, print para diagnóstico de timing/foco."""
    if not _diagnostico_ui_ativo():
        return
    try:
        titulos = bot.windows.Janela.titulos_janelas()
        bot.logger.informar(f"[diag-ui] {contexto} | titulos={titulos}")
    except Exception as e:
        bot.logger.informar(f"[diag-ui] {contexto} | falha ao listar títulos: {e}")

    if capturar_print:
        caminho = _capturar_print_erro("diag_ui")
        if caminho:
            bot.logger.informar(f"[diag-ui] print={caminho}")


def _encontrar_imagem(caminho: str, confianca: float = 0.8, segundos: int = 5):
    """Retorna coordenada da imagem se encontrada, None caso contrário."""
    try:
        return bot.imagem.procurar_imagem(caminho, confianca=confianca, segundos=segundos)
    except Exception:
        return None


def _extrair_xy_centro(coord) -> tuple[int, int] | None:
    """Extrai (x, y) do centro a partir de Coordenada (left, top, width, height) ou similar."""
    try:
        if hasattr(coord, "__len__") and len(coord) >= 2:
            x, y = int(coord[0]), int(coord[1])
            if len(coord) >= 4:
                x += int(coord[2]) // 2
                y += int(coord[3]) // 2
            return (x, y)
        left = getattr(coord, "left", None) or getattr(coord, "x", None)
        top = getattr(coord, "top", None) or getattr(coord, "y", None)
        w = getattr(coord, "width", None) or 0
        h = getattr(coord, "height", None) or 0
        if left is not None and top is not None:
            x = int(left + w / 2) if w else int(left)
            y = int(top + h / 2) if h else int(top)
            return (x, y)
    except (TypeError, ValueError, IndexError):
        pass
    return None


def _titulos_janelas_seguro() -> list[str]:
    """Obtém títulos de janelas sem derrubar o fluxo em casos de OpenProcess/Acesso negado."""
    try:
        titulos = bot.windows.Janela.titulos_janelas()
        if isinstance(titulos, (list, tuple, set)):
            return [str(t).strip() for t in titulos if str(t).strip()]
        if titulos:
            ts = str(titulos).strip()
            return [ts] if ts else []
    except Exception as e:
        bot.logger.informar(f"Falha ao obter títulos via bot.windows: {e}")

    try:
        from pywinauto import Desktop

        saida = []
        for w in Desktop(backend="win32").windows():
            try:
                tt = (w.window_text() or "").strip()
                if tt:
                    saida.append(tt)
            except Exception:
                continue
        return saida
    except Exception:
        return []


def _obter_pausa_interacao_nbs_s() -> float:
    """Retorna pausa entre interações para reduzir corrida de foco/timing em testes."""
    try:
        v = bot.configfile.obter_opcao_ou("nbs", "espera_interacao_ms")
        if v is not None and str(v).strip() != "":
            ms = int(str(v).strip())
            return max(0.05, min(2.0, ms / 1000.0))
    except Exception:
        pass
    # Padrão mais conservador para ambiente de homologação/VM.
    return 0.35


def _pausa_interacao_nbs() -> None:
    sleep(_obter_pausa_interacao_nbs_s())


def _focar_janela_nbs_relevante(preferir_entradas: bool = True) -> bool:
    """Foca uma janela relevante do NBS para evitar teclado/clique em apps externos."""
    try:
        if preferir_entradas:
            w_ent = _obter_janela_nota_fiscal_diversas()
            if w_ent is not None:
                try:
                    w_ent.set_focus()
                    _pausa_interacao_nbs()
                    return True
                except Exception:
                    pass
    except Exception:
        pass

    prioridades = (
        "entrada diversas",
        "nbs-nota fiscal diversas",
        "nota fiscal diversas",
        "sistema fiscal",
        "empresa/filial",
        "exercício contábil",
        "nbs",
    )

    titulos = _titulos_janelas_seguro()
    for chave in prioridades:
        for t in titulos:
            tl = str(t).lower()
            if chave in tl:
                try:
                    bot.windows.Janela(t).focar()
                    _pausa_interacao_nbs()
                    return True
                except Exception:
                    continue

    try:
        from pywinauto import Desktop
        for chave in prioridades:
            for w in Desktop(backend="win32").windows():
                try:
                    tt = (w.window_text() or "").strip().lower()
                    if tt and chave in tt:
                        w.set_focus()
                        _pausa_interacao_nbs()
                        return True
                except Exception:
                    continue
    except Exception:
        pass

    return False


def _aguardar_contexto_nbs(
    contexto: str,
    timeout_s: float = 10.0,
    preferir_entradas: bool = True,
    ciclos_estaveis: int = 2,
) -> bool:
    """Aguarda janelas NBS relevantes e tenta focar explicitamente antes de interagir."""
    total = max(1, int((timeout_s * 1000) // 250))
    estavel = 0
    for _ in range(total):
        titulos = [t.lower() for t in _titulos_janelas_seguro()]
        ok = any(
            k in " | ".join(titulos)
            for k in ("sistema fiscal", "nbs-nota fiscal diversas", "entrada diversas", "nota fiscal diversas", "empresa/filial")
        )
        if ok and _focar_janela_nbs_relevante(preferir_entradas=preferir_entradas):
            estavel += 1
            if estavel >= max(1, int(ciclos_estaveis)):
                return True
        else:
            estavel = 0
        sleep(0.25)

    bot.logger.alertar(f"Contexto NBS não estabilizou a tempo: {contexto}")
    return False


def _digitar_com_foco_nbs(texto: str, contexto: str, preferir_entradas: bool = True) -> bool:
    """Digita texto somente após confirmar foco em janela NBS relevante."""
    if not _aguardar_contexto_nbs(contexto, timeout_s=8.0, preferir_entradas=preferir_entradas):
        return False
    bot.teclado.digitar_teclado(texto)
    _pausa_interacao_nbs()
    return True


def _tecla_com_foco_nbs(tecla: str, contexto: str, preferir_entradas: bool = True) -> bool:
    """Pressiona tecla somente após confirmar foco em janela NBS relevante."""
    if not _aguardar_contexto_nbs(contexto, timeout_s=8.0, preferir_entradas=preferir_entradas):
        return False
    bot.teclado.apertar_tecla(tecla)
    _pausa_interacao_nbs()
    return True


def _aguardar_sistema_fiscal_pronto(timeout_s: float = 12.0) -> bool:
    """Aguarda Sistema Fiscal (ou já estar em Entradas) em estado estável para seguir o fluxo."""
    total = max(1, int((timeout_s * 1000) // 250))
    estavel = 0
    for _ in range(total):
        titulos = _titulos_janelas_seguro()
        tem_sistema = any("sistema fiscal" in t.lower() for t in titulos)
        tem_entradas = _obter_janela_nota_fiscal_diversas() is not None
        if (tem_sistema or tem_entradas) and not _janela_empresa_filial_aberta():
            estavel += 1
            if estavel >= 2:
                return True
        else:
            estavel = 0
        sleep(0.25)
    return False


def _aguardar_entradas_pronta_para_incluir(timeout_s: float = 8.0) -> bool:
    """Aguarda a janela Entradas/Nota Fiscal Diversas estar pronta antes do clique em Incluir."""
    total = max(1, int((timeout_s * 1000) // 250))
    estavel = 0
    for _ in range(total):
        w = _obter_janela_nota_fiscal_diversas()
        if w is not None:
            try:
                w.set_focus()
            except Exception:
                pass
            try:
                tt = (w.window_text() or "").strip().lower()
            except Exception:
                tt = ""
            if any(k in tt for k in ("entradas", "nota fiscal diversa", "nota fiscal diversas")):
                estavel += 1
                if estavel >= 2:
                    return True
            else:
                estavel = 0
        else:
            estavel = 0
        sleep(0.25)
    return False


def _obter_xy_nota_fiscal_entradas() -> tuple[int, int]:
    """Obtém a coordenada do botão Nota Fiscal Entradas (configurável via .ini)."""
    try:
        xy_cfg = bot.configfile.obter_opcao_ou("nbs", "nota_fiscal_entradas_xy")
        if xy_cfg:
            nums = [int(n) for n in re.findall(r"-?\d+", str(xy_cfg))]
            if len(nums) >= 2:
                return (nums[0], nums[1])
    except Exception:
        pass

    try:
        x_cfg = bot.configfile.obter_opcao_ou("nbs", "nota_fiscal_entradas_x")
        y_cfg = bot.configfile.obter_opcao_ou("nbs", "nota_fiscal_entradas_y")
        if x_cfg is not None and y_cfg is not None:
            return (int(str(x_cfg).strip()), int(str(y_cfg).strip()))
    except Exception:
        pass

    return NOTA_FISCAL_ENTRADAS_XY


def _clicar_nota_fiscal_entradas_por_coordenada() -> bool:
    """Fallback: clica no botão Nota Fiscal Entradas por coordenada global."""
    try:
        import pyautogui

        if not _aguardar_contexto_nbs("click coordenada Nota Fiscal Entradas", timeout_s=8.0, preferir_entradas=False):
            return False

        x, y = _obter_xy_nota_fiscal_entradas()
        pyautogui.moveTo(x, y, duration=0.15)
        _pausa_interacao_nbs()
        pyautogui.click()
        bot.logger.informar(f"Nota fiscal entradas via coordenada ({x},{y})")
        return True
    except Exception as e:
        bot.logger.informar(f"Clique por coordenada (Entradas) falhou: {e}")
        return False


def _obter_xy_incluir_entrada_diversa() -> tuple[int, int]:
    """Obtém a coordenada do botão Incluir Entrada (Só Diversa) (configurável via .ini)."""
    try:
        xy_cfg = bot.configfile.obter_opcao_ou("nbs", "incluir_entrada_diversa_xy")
        if xy_cfg:
            nums = [int(n) for n in re.findall(r"-?\d+", str(xy_cfg))]
            if len(nums) >= 2:
                return (nums[0], nums[1])
    except Exception:
        pass

    try:
        x_cfg = bot.configfile.obter_opcao_ou("nbs", "incluir_entrada_diversa_x")
        y_cfg = bot.configfile.obter_opcao_ou("nbs", "incluir_entrada_diversa_y")
        if x_cfg is not None and y_cfg is not None:
            return (int(str(x_cfg).strip()), int(str(y_cfg).strip()))
    except Exception:
        pass

    return INCLUIR_ENTRADA_DIVERSA_XY


def _clicar_incluir_entrada_diversa_por_coordenada() -> bool:
    """Fallback: clica no botão Incluir Entrada (Só Diversa) por coordenada global."""
    try:
        import pyautogui

        if not _aguardar_entradas_pronta_para_incluir(timeout_s=4.0):
            bot.logger.alertar("Clique por coordenada (Incluir Entrada) abortado: janela Entradas ainda não está pronta")
            return False
        if not _aguardar_contexto_nbs("click coordenada Incluir Entrada", timeout_s=8.0, preferir_entradas=True):
            return False

        x, y = _obter_xy_incluir_entrada_diversa()
        pyautogui.moveTo(x, y, duration=0.15)
        _pausa_interacao_nbs()
        pyautogui.click()
        bot.logger.informar(f"Incluir Entrada (Só Diversa) via coordenada ({x},{y})")
        return True
    except Exception as e:
        bot.logger.informar(f"Clique por coordenada (Incluir Entrada) falhou: {e}")
        return False


def _clicar_por_imagem_se_existir(caminho: str, confianca: float = 0.8, segundos: int = 5) -> bool:
    """Fallback: clica por reconhecimento de imagem se o arquivo existir e for encontrado.
    Move o mouse explicitamente para as coordenadas antes do clique (evita clicar na posição atual)."""
    _dispensar_popup_data_hora_invalida(tentativas=2)
    _dispensar_popup_listagem_notas(tentativas=2)
    _dispensar_janela_notas_ativas_sair(tentativas=2)
    img = _encontrar_imagem(caminho, confianca, segundos)
    if img:
        xy = _extrair_xy_centro(img)
        if xy:
            import pyautogui
            pyautogui.moveTo(xy[0], xy[1], duration=0.15)
            sleep(0.1)
            pyautogui.click()
        else:
            return False
        bot.logger.informar(f"Fallback imagem: {caminho}")
        return True
    return False


def _ativar_aba_por_imagem_robusta(
    chave_imagem: str,
    nome_aba: str,
    tentativas: int = 3,
    confianca: float = 0.85,
    segundos_busca: int = 3,
    validador: Callable[[], bool] | None = None,
) -> bool:
    """Ativa aba por imagem com retry curto e validação opcional pós-clique."""
    caminho = IMAGENS_FALLBACK.get(chave_imagem)
    if not caminho:
        return False

    total = max(1, int(tentativas))
    for tentativa in range(1, total + 1):
        clicou = _clicar_por_imagem_se_existir(caminho, confianca=confianca, segundos=segundos_busca)
        if not clicou:
            sleep(0.25)
            continue

        sleep(0.45)
        if not validador:
            bot.logger.informar(f"Aba {nome_aba}: clique por imagem (tentativa {tentativa}/{total})")
            return True

        try:
            if validador():
                bot.logger.informar(
                    f"Aba {nome_aba}: clique por imagem validado (tentativa {tentativa}/{total})"
                )
                return True
        except Exception:
            pass

        bot.logger.alertar(
            f"Aba {nome_aba}: clique por imagem sem validação de contexto (tentativa {tentativa}/{total})"
        )

    return False


def _obter_janela_nbs_segura(titulo_parcial: str = "NBS"):
    """Obtém janela NBS de forma segura, desambiguando se houver múltiplas com mesmo título.
    Retorna a primeira janela que contenha titulo_parcial no título, ou None."""
    try:
        titulos = bot.windows.Janela.titulos_janelas()
        if not titulos:
            return None
        candidatas = [t for t in (list(titulos) if isinstance(titulos, (list, tuple)) else [titulos]) if t and titulo_parcial.lower() in str(t).lower()]
        if candidatas:
            return bot.windows.Janela(candidatas[0])
    except Exception:
        pass
    # Fallback: tentar janela ativa
    try:
        return bot.windows.Janela()
    except Exception:
        return None

def _janela_tem_texto_popup_data_hora(janela) -> bool:
    """Valida se a janela contém o erro clássico de data/hora inválida do NBS."""
    try:
        titulo = (janela.window_text() or "").strip().lower()
        if "nbs" not in titulo:
            return False
        textos = [titulo]
        for elem in janela.descendants():
            try:
                txt = (elem.window_text() or "").strip().lower()
                if txt:
                    textos.append(txt)
            except Exception:
                continue
        conteudo = " | ".join(textos)
        return (
            "not a valid date and time" in conteudo
            or ("data" in conteudo and "hora" in conteudo and "inv" in conteudo)
        )
    except Exception:
        return False


def _listar_janelas_popup_data_hora() -> list:
    """Lista janelas top-level que representam popup modal de data/hora inválida."""
    try:
        from pywinauto import Desktop
        janelas = []
        for janela in Desktop(backend="win32").windows():
            if _janela_tem_texto_popup_data_hora(janela):
                janelas.append(janela)
        return janelas
    except Exception:
        return []


def _capturar_print_erro(prefixo: str) -> str | None:
    """Captura screenshot para diagnóstico e retorna o caminho salvo."""
    try:
        import pyautogui
        pasta = Path("imagens") / "erros"
        pasta.mkdir(parents=True, exist_ok=True)
        nome = f"{prefixo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        caminho = pasta / nome
        pyautogui.screenshot(str(caminho))
        return str(caminho)
    except Exception:
        return None


def _dispensar_popup_data_hora_invalida(tentativas: int = 3) -> bool:
    """Fecha popup de data/hora inválida (modal/top-level) com clique em OK ou Enter."""
    fechou = False
    for _ in range(max(1, tentativas)):
        janelas_popup = _listar_janelas_popup_data_hora()
        if not janelas_popup:
            break

        for janela_popup in janelas_popup:
            tentou_fechar = False

            # Estratégia 1: botão OK na própria janela modal
            for _cls_btn in ("TBitBtn", "TButton", "Button"):
                try:
                    for btn in janela_popup.descendants(class_name=_cls_btn):
                        try:
                            btxt = (btn.window_text() or "").strip().lower()
                            if btxt in ("ok", "&ok"):
                                btn.click_input()
                                tentou_fechar = True
                                fechou = True
                                bot.logger.alertar("Popup data/hora inválida fechado (OK)")
                                break
                        except Exception:
                            continue
                except Exception:
                    continue
                if tentou_fechar:
                    break

            # Estratégia 2: Enter na janela do popup
            if not tentou_fechar:
                try:
                    janela_popup.set_focus()
                    janela_popup.type_keys("{ENTER}")
                    fechou = True
                    bot.logger.alertar("Popup data/hora inválida fechado (ENTER)")
                    tentou_fechar = True
                except Exception:
                    pass

            # Estratégia 3: Enter global como fallback final
            if not tentou_fechar:
                try:
                    bot.teclado.apertar_tecla("enter")
                    fechou = True
                    bot.logger.alertar("Popup data/hora inválida fechado (ENTER global)")
                except Exception:
                    pass

        sleep(0.2)

    return fechou


def _ha_popup_data_hora_invalida() -> bool:
    """Detecta popup de data/hora inválida considerando janelas modais top-level."""
    return bool(_listar_janelas_popup_data_hora())


def _garantir_sem_popup_data_hora(contexto: str, ciclos: int = 10) -> None:
    """Garante que popup de data/hora foi fechado antes de continuar o fluxo."""
    total = max(1, int(ciclos))
    for _ in range(total):
        _dispensar_popup_data_hora_invalida(tentativas=3)
        if not _ha_popup_data_hora_invalida():
            return
        sleep(0.25)
    caminho_print = _capturar_print_erro("popup_data_hora")
    msg_print = f" Print: {caminho_print}" if caminho_print else ""
    raise ErroNBS(f"Popup de data/hora inválida permaneceu aberto ({contexto}).{msg_print}")


def _janela_e_listagem_notas_com_ok(janela) -> bool:
    """Identifica janela de listagem/notas com botão OK que bloqueia o fluxo."""
    try:
        titulo = (janela.window_text() or "").strip().lower()
        if not titulo:
            return False
        if not any(k in titulo for k in ("nota", "notas", "entrada", "entradas", "diversa", "fiscal", "nbs")):
            return False

        tem_ok = False
        for cls_btn in ("TBitBtn", "TButton", "Button"):
            for btn in janela.descendants(class_name=cls_btn):
                try:
                    txt = (btn.window_text() or "").strip().lower()
                    if txt in ("ok", "&ok"):
                        tem_ok = True
                        break
                except Exception:
                    continue
            if tem_ok:
                break
        if not tem_ok:
            return False

        for cls_grid in ("TDBGrid", "TStringGrid", "TListView", "TListBox", "SysListView32", "ListBox"):
            try:
                if janela.descendants(class_name=cls_grid):
                    return True
            except Exception:
                continue

        return "nota" in titulo or "entrada" in titulo
    except Exception:
        return False


def _dispensar_popup_listagem_notas(tentativas: int = 3) -> bool:
    """Fecha janela de listagem/notas com botão OK quando ela aparece antes de Entradas."""
    try:
        from pywinauto import Desktop
    except Exception:
        return False

    fechou = False
    for _ in range(max(1, tentativas)):
        alguma = False
        try:
            for w in Desktop(backend="win32").windows():
                try:
                    if not _janela_e_listagem_notas_com_ok(w):
                        continue
                    alguma = True
                    for cls_btn in ("TBitBtn", "TButton", "Button"):
                        clicou = False
                        for btn in w.descendants(class_name=cls_btn):
                            try:
                                txt = (btn.window_text() or "").strip().lower()
                                if txt in ("ok", "&ok"):
                                    btn.click_input()
                                    bot.logger.alertar("Janela de listagem/notas fechada (OK)")
                                    sleep(0.25)
                                    clicou = True
                                    fechou = True
                                    break
                            except Exception:
                                continue
                        if clicou:
                            break
                except Exception:
                    continue
        except Exception:
            pass

        if not alguma:
            break
        sleep(0.2)

    return fechou


def _janela_notas_ativas_com_sair(janela) -> bool:
    """Identifica a janela de listagem de notas que exibe botão Sair."""
    try:
        titulo = (janela.window_text() or "").strip().lower()
        if not titulo:
            return False

        # Exemplo real: "Notas Fiscais Ativas, Denegadas, Canceladas e Rejeitadas"
        if "notas fiscais" not in titulo:
            return False
        if not any(k in titulo for k in ("ativas", "denegadas", "canceladas", "rejeitadas")):
            return False

        for cls_btn in ("TBitBtn", "TButton", "Button"):
            for btn in janela.descendants(class_name=cls_btn):
                try:
                    txt = (btn.window_text() or "").strip().lower()
                    if txt in ("sair", "&sair"):
                        return True
                except Exception:
                    continue
    except Exception:
        return False
    return False


def _dispensar_janela_notas_ativas_sair(tentativas: int = 3) -> bool:
    """Fecha a janela 'Notas Fiscais ...' clicando no botão Sair quando ela estiver aberta."""
    try:
        from pywinauto import Desktop
    except Exception:
        return False

    fechou = False
    for _ in range(max(1, tentativas)):
        alguma = False
        try:
            for w in Desktop(backend="win32").windows():
                try:
                    if not _janela_notas_ativas_com_sair(w):
                        continue
                    alguma = True
                    clicou = False
                    for cls_btn in ("TBitBtn", "TButton", "Button"):
                        for btn in w.descendants(class_name=cls_btn):
                            try:
                                txt = (btn.window_text() or "").strip().lower()
                                if txt in ("sair", "&sair"):
                                    btn.click_input()
                                    bot.logger.alertar("Janela de Notas Fiscais fechada (Sair)")
                                    sleep(0.25)
                                    clicou = True
                                    fechou = True
                                    break
                            except Exception:
                                continue
                        if clicou:
                            break
                except Exception:
                    continue
        except Exception:
            pass

        if not alguma:
            break
        sleep(0.2)

    return fechou


def _obter_janela_empresa_filial():
    """Retorna a janela modal Empresa/Filial/Exercício Contábil quando estiver aberta."""
    try:
        from pywinauto import Desktop
        for w in Desktop(backend="win32").windows():
            try:
                titulo = (w.window_text() or "").strip().lower()
                if (
                    "empresa/filial" in titulo
                    or ("empresa" in titulo and "filial" in titulo)
                    or "exercício contábil" in titulo
                    or "exercicio contabil" in titulo
                    or "exercício" in titulo
                    or "exercicio" in titulo
                ):
                    return w
            except Exception:
                continue
    except Exception:
        pass
    return None


def _janela_empresa_filial_aberta() -> bool:
    """Informa se a janela Empresa/Filial/Exercício ainda está aberta."""
    return _obter_janela_empresa_filial() is not None


def _aguardar_fechamento_empresa_filial(timeout_s: float = 6.0) -> bool:
    """Aguarda fechamento da janela Empresa/Filial após confirmar."""
    total = max(1, int((timeout_s * 1000) // 250))
    for _ in range(total):
        if not _janela_empresa_filial_aberta():
            return True
        sleep(0.25)
    return not _janela_empresa_filial_aberta()


def _preencher_empresa_filial_e_confirmar(cod_empresa: str, cod_filial: str) -> bool:
    """Preenche Empresa/Filial na janela modal e clica em Confirmar."""
    janela = _obter_janela_empresa_filial()
    if not janela:
        return False

    try:
        janela.set_focus()
        sleep(0.15)
    except Exception:
        pass

    try:
        # Prioridade 1: dropdowns/combos (como na janela Empresa/Filial do NBS)
        combos = []
        for cls_combo in ("TwwDBLookupCombo", "TDBLookupComboBox", "TComboBox", "TwwDBComboBox"):
            for cb in janela.descendants(class_name=cls_combo):
                try:
                    if hasattr(cb, "is_visible") and not cb.is_visible():
                        continue
                    r = cb.rectangle()
                    combos.append((r.top, r.left, cb))
                except Exception:
                    continue
        combos.sort(key=lambda x: (x[0], x[1]))

        preencheu = False
        if len(combos) >= 2:
            for valor, (_, _, combo) in ((str(cod_empresa), combos[0]), (str(cod_filial), combos[1])):
                try:
                    combo.click_input()
                    sleep(0.1)
                    combo.type_keys("^a" + valor, with_spaces=True)
                    sleep(0.1)
                    combo.type_keys("{TAB}", with_spaces=False)
                except Exception:
                    bot.teclado.digitar_teclado(valor)
                    bot.teclado.apertar_tecla("tab")
                sleep(0.2)
            preencheu = True

        # Prioridade 2: campos editáveis (fallback)
        if not preencheu:
            campos = []
            for cls in ("TOvcPictureField", "TOvcDbPictureField", "TEdit", "TDBEdit"):
                for campo in janela.descendants(class_name=cls):
                    try:
                        if hasattr(campo, "is_visible") and not campo.is_visible():
                            continue
                        r = campo.rectangle()
                        campos.append((r.top, r.left, campo))
                    except Exception:
                        continue
            campos.sort(key=lambda x: (x[0], x[1]))

            if len(campos) >= 2:
                for valor, (_, _, campo) in ((str(cod_empresa), campos[0]), (str(cod_filial), campos[1])):
                    try:
                        campo.click_input()
                        sleep(0.1)
                        campo.type_keys("^a{BACKSPACE}" + valor, with_spaces=False)
                    except Exception:
                        bot.teclado.digitar_teclado(valor)
                    sleep(0.15)
                preencheu = True

        # Prioridade 3: fallback de teclado direto na modal
        if not preencheu:
            try:
                janela.set_focus()
                sleep(0.1)
                bot.teclado.digitar_teclado(str(cod_empresa))
                bot.teclado.apertar_tecla("tab")
                sleep(0.1)
                bot.teclado.digitar_teclado(str(cod_filial))
                bot.teclado.apertar_tecla("tab")
                preencheu = True
                bot.logger.informar("Empresa/Filial preenchida via fallback de teclado")
            except Exception:
                preencheu = False

        if not preencheu:
            return False

        confirmou = False
        for cls_btn in ("TBitBtn", "TButton", "Button"):
            for btn in janela.descendants(class_name=cls_btn):
                try:
                    txt = (btn.window_text() or "").strip().lower()
                    if "confirma" in txt:
                        btn.click_input()
                        confirmou = True
                        break
                except Exception:
                    continue
            if confirmou:
                break

        if not confirmou:
            try:
                janela.set_focus()
                janela.type_keys("{ENTER}")
                confirmou = True
            except Exception:
                confirmou = False

        if not confirmou:
            return False

        if not _aguardar_fechamento_empresa_filial(timeout_s=6.0):
            bot.logger.alertar("Janela Empresa/Filial permaneceu aberta após tentativa de Confirmar")
            return False

        sleep(0.3)
        return True
    except Exception:
        return False


def _resolver_empresa_filial_com_retry(
    cod_empresa: str,
    cod_filial: str,
    tentativas: int = 3,
) -> bool:
    """Resolve modal Empresa/Filial com múltiplas tentativas antes de avançar o fluxo."""
    total = max(1, int(tentativas))
    for tentativa in range(1, total + 1):
        janela = _obter_janela_empresa_filial()
        if not janela:
            return True

        bot.logger.informar(
            f"Tentando confirmar Empresa/Filial ({tentativa}/{total})"
        )
        if _preencher_empresa_filial_e_confirmar(cod_empresa, cod_filial):
            if not _janela_empresa_filial_aberta():
                return True

        # Fallback extra: Enter na própria modal pode confirmar quando botão não expõe texto.
        try:
            janela = _obter_janela_empresa_filial()
            if janela:
                janela.set_focus()
                janela.type_keys("{ENTER}")
                sleep(0.5)
                if not _janela_empresa_filial_aberta():
                    bot.logger.informar("Empresa/Filial confirmada via Enter na modal")
                    return True
        except Exception:
            pass

        sleep(0.3)

    return not _janela_empresa_filial_aberta()


def _garantir_empresa_filial_resolvida_antes_entradas(
    cod_empresa: str,
    cod_filial: str,
    timeout_aparecer_s: float = 5.0,
) -> bool:
    """Garante que uma modal tardia de Empresa/Filial seja resolvida antes de clicar em Entradas."""
    total = max(1, int((timeout_aparecer_s * 1000) // 250))
    apareceu = False
    for _ in range(total):
        if _janela_empresa_filial_aberta():
            apareceu = True
            break
        sleep(0.25)

    if not apareceu:
        return True

    bot.logger.alertar("Janela Empresa/Filial detectada tardiamente. Tentando resolver antes de Entradas.")
    return _resolver_empresa_filial_com_retry(cod_empresa, cod_filial, tentativas=4)


def _obter_janela_nota_fiscal_diversas():
    """Retorna a janela de Entradas (intermediária) ou Nota Fiscal Diversas quando aberta."""
    try:
        from pywinauto import Desktop
        for w in Desktop(backend="win32").windows():
            try:
                titulo = (w.window_text() or "").strip().lower()
                if not titulo:
                    continue
                if (
                    titulo == "entradas"
                    or titulo.startswith("entradas ")
                    or
                    "nbs-nota fiscal diversas" in titulo
                    or "nota fiscal diversas" in titulo
                    or "nota fiscal diversa" in titulo
                ):
                    return w
            except Exception:
                continue
    except Exception:
        pass
    return None


def _aguardar_janela_nota_fiscal_diversas(timeout_s: float = 8.0) -> bool:
    """Aguarda a abertura da janela de Entradas/Nota Fiscal Diversas."""
    total = max(1, int((timeout_s * 1000) // 250))
    for _ in range(total):
        if _obter_janela_nota_fiscal_diversas() is not None:
            return True
        sleep(0.25)
    return _obter_janela_nota_fiscal_diversas() is not None


def _obter_janela_entrada_diversas_operacao():
    """Retorna a janela principal 'Entrada Diversas / Operação...' quando aberta."""
    try:
        from pywinauto import Desktop
        for w in Desktop(backend="win32").windows():
            try:
                titulo = (w.window_text() or "").strip().lower()
                if not titulo:
                    continue
                if "entrada diversas" in titulo and ("opera" in titulo or "operação" in titulo):
                    return w
            except Exception:
                continue
    except Exception:
        pass
    return None


def _preencher_capa_por_labels(janela_entrada, processo: str, data_emissao_digits: str) -> bool:
    """Preenche Nº NF, Série, Emissão e Entrada por rótulo, evitando dependência de foco/Tab."""
    try:
        _labels = []
        for _lbl in janela_entrada.elementos(class_name="TLabel", top_level_only=False):
            try:
                _txt = (_lbl.window_text() or "").strip().lower()
                if _txt:
                    _labels.append((_txt, _lbl.rectangle(), _lbl))
            except Exception:
                continue

        _campos = []
        for _cls in ("TOvcDbPictureField", "TOvcPictureField", "TDBEdit", "TEdit"):
            for _f in janela_entrada.elementos(class_name=_cls, top_level_only=False):
                try:
                    _r = _f.rectangle()
                    _campos.append((_r.top, _r.left, _r, _f))
                except Exception:
                    continue
        if not _campos:
            return False

        def _achar_campo(*keywords: str):
            for _txt, _lr, _lbl in _labels:
                if any(k in _txt for k in keywords):
                    _cands = [
                        (abs(_r.top - _lr.top), _r.left, _f)
                        for _, _, _r, _f in _campos
                        if _r.left >= _lr.left and abs(_r.top - _lr.top) <= 28
                    ]
                    if _cands:
                        _cands.sort(key=lambda x: (x[0], x[1]))
                        return _cands[0][2]
            return None

        _campo_nf = _achar_campo("número", "numero", "nota fiscal", "n.f")
        _campo_serie = _achar_campo("série", "serie")
        _campo_emissao = _achar_campo("emissão", "emissao")
        _campo_entrada = _achar_campo("entrada")
        if not all((_campo_nf, _campo_serie, _campo_emissao, _campo_entrada)):
            return False

        for _campo, _valor in (
            (_campo_nf, str(processo)),
            (_campo_serie, "E"),
            (_campo_emissao, str(data_emissao_digits)),
            (_campo_entrada, str(data_emissao_digits)),
        ):
            _campo.click_input()
            sleep(0.1)
            try:
                _campo.type_keys("^a{BACKSPACE}" + _valor, with_spaces=False)
            except Exception:
                _campo.type_keys("^a" + _valor, with_spaces=False)
            sleep(0.15)

        bot.logger.informar("Capa preenchida por labels (Nº NF, Série, Emissão, Entrada)")
        return True
    except Exception as _e:
        bot.logger.informar(f"Preenchimento por labels falhou: {_e}")
        return False


def _clicar_nota_fiscal_entradas_robusto(
    tentativas: int = 3,
    cod_empresa: str | None = None,
    cod_filial: str | None = None,
) -> bool:
    """Tenta acessar Nota Fiscal Entradas por imagem/menu e valida a abertura da tela de Diversas."""
    def _entrada_foi_aberta() -> bool:
        return _obter_janela_nota_fiscal_diversas() is not None

    def _focar_sistema_fiscal() -> bool:
        try:
            titulos = bot.windows.Janela.titulos_janelas()
            lista_titulos = list(titulos) if isinstance(titulos, (list, tuple, set)) else ([titulos] if titulos else [])
            for t in lista_titulos:
                ts = str(t or "")
                if ts and "sistema fiscal" in ts.lower():
                    bot.windows.Janela(ts).focar()
                    sleep(0.2)
                    return True
        except Exception:
            pass
        try:
            from pywinauto import Desktop
            for w in Desktop(backend="win32").windows():
                try:
                    tt = (w.window_text() or "").strip().lower()
                    if "sistema fiscal" in tt:
                        w.set_focus()
                        sleep(0.2)
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

    total = max(1, int(tentativas))
    for tentativa in range(1, total + 1):
        _dispensar_janela_notas_ativas_sair(tentativas=2)
        _focar_sistema_fiscal()
        if _janela_empresa_filial_aberta():
            _diagnostico_ui_snapshot(
                f"empresa_filial_aberta_antes_entradas_tentativa_{tentativa}",
                capturar_print=True,
            )
            if cod_empresa and cod_filial:
                bot.logger.alertar(
                    f"Empresa/Filial ainda aberta antes de Entradas (tentativa {tentativa}/{total}). Tentando resolver novamente."
                )
                if not _resolver_empresa_filial_com_retry(str(cod_empresa), str(cod_filial), tentativas=3):
                    sleep(0.3)
                    continue
                if _janela_empresa_filial_aberta():
                    sleep(0.3)
                    continue
            else:
                bot.logger.alertar(
                    f"Empresa/Filial ainda aberta antes de Entradas (tentativa {tentativa}/{total}). Bloqueando clique global."
                )
                return False

        _dispensar_popup_listagem_notas(tentativas=3)
        _dispensar_popup_data_hora_invalida(tentativas=4)
        if _ha_popup_data_hora_invalida():
            bot.logger.alertar(
                f"Popup de data/hora ainda aberto antes de clicar Entradas (tentativa {tentativa}/{total})"
            )
            sleep(0.3)
            continue

        # 1) Imagem dedicada
        if _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["nota_fiscal_entradas"], confianca=0.9, segundos=4):
            if _aguardar_janela_nota_fiscal_diversas(timeout_s=4.0):
                try:
                    _jnf = _obter_janela_nota_fiscal_diversas()
                    if _jnf:
                        _jnf.set_focus()
                except Exception:
                    pass
                bot.logger.informar(f"Nota fiscal entradas via botao do Sistema Fiscal (imagem) (tentativa {tentativa}/{total})")
                return True
            bot.logger.alertar(
                f"Clique em Nota Fiscal Entradas (imagem) não abriu NBS-Nota Fiscal Diversas (tentativa {tentativa}/{total})"
            )

        # 1.1) Fallback por coordenada fixa (informada por captura em VM)
        if _clicar_nota_fiscal_entradas_por_coordenada():
            if _aguardar_janela_nota_fiscal_diversas(timeout_s=4.0):
                try:
                    _jnf = _obter_janela_nota_fiscal_diversas()
                    if _jnf:
                        _jnf.set_focus()
                except Exception:
                    pass
                bot.logger.informar(
                    f"Nota fiscal entradas via coordenada (tentativa {tentativa}/{total})"
                )
                return True
            bot.logger.alertar(
                f"Clique em Nota Fiscal Entradas (coordenada) não abriu NBS-Nota Fiscal Diversas (tentativa {tentativa}/{total})"
            )

        # 2) Fallback por menu Notas Fiscais -> Nota Fiscal Entradas
        try:
            from pywinauto import Desktop
            for w in Desktop(backend="win32").windows():
                try:
                    titulo = (w.window_text() or "")
                    if "sistema fiscal" not in titulo.lower() and "nbs" not in titulo.lower():
                        continue
                    menu_notas = None
                    for item in w.descendants(class_name="TMenuItem"):
                        txt = (item.window_text() or "").strip().lower()
                        if "notas fiscais" in txt:
                            menu_notas = item
                            break
                    if menu_notas is None:
                        continue
                    menu_notas.click_input()
                    sleep(0.2)
                    for item in w.descendants(class_name="TMenuItem"):
                        txt = (item.window_text() or "").strip().lower()
                        if "entrada" in txt:
                            item.click_input()
                            if _entrada_foi_aberta() and _aguardar_janela_nota_fiscal_diversas(timeout_s=4.0):
                                try:
                                    _jnf = _obter_janela_nota_fiscal_diversas()
                                    if _jnf:
                                        _jnf.set_focus()
                                except Exception:
                                    pass
                                bot.logger.informar(
                                    f"Nota fiscal entradas via menu Notas Fiscais (tentativa {tentativa}/{total})"
                                )
                                return True
                            break
                except Exception:
                    continue
        except Exception:
            pass

        # 3) Fallback teclado (atalho legado do menu)
        try:
            _focar_sistema_fiscal()
            if not _tecla_com_foco_nbs("alt", "fallback teclado Entradas - ALT", preferir_entradas=False):
                continue
            if not _digitar_com_foco_nbs("e", "fallback teclado Entradas - E", preferir_entradas=False):
                continue
            if not _tecla_com_foco_nbs("enter", "fallback teclado Entradas - ENTER", preferir_entradas=False):
                continue
            if _entrada_foi_aberta() and _aguardar_janela_nota_fiscal_diversas(timeout_s=4.0):
                try:
                    _jnf = _obter_janela_nota_fiscal_diversas()
                    if _jnf:
                        _jnf.set_focus()
                except Exception:
                    pass
                bot.logger.informar(f"Nota fiscal entradas via teclado/menu (tentativa {tentativa}/{total})")
                return True
        except Exception:
            pass

    return False


def _normalizar_data_ddmmyyyy(valor) -> str | None:
    """Normaliza datas em formatos comuns para DD/MM/YYYY."""
    if valor is None:
        return None
    s = str(valor).strip()
    if not s:
        return None

    # ISO (ex.: 2026-03-16T16:16:00Z)
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


def _obter_data_fechamento_sistema_fiscal() -> str | None:
    """Lê a data de fechamento exibida no Sistema Fiscal (ex.: 'Fechamento: 31/10/2020')."""
    try:
        from pywinauto import Desktop
        padrao = re.compile(r"fechamento\s*:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
        for w in Desktop(backend="win32").windows():
            try:
                titulo = (w.window_text() or "").strip().lower()
                if "sistema fiscal" not in titulo and "nbs" not in titulo:
                    continue

                # Verifica texto da própria janela e de todos os descendentes.
                textos = [w.window_text() or ""]
                for elem in w.descendants():
                    try:
                        txt = elem.window_text() or ""
                        if txt:
                            textos.append(txt)
                    except Exception:
                        continue

                blob = " | ".join(textos)
                m = padrao.search(blob)
                if m:
                    return m.group(1)
            except Exception:
                continue
    except Exception:
        pass
    return None


def _validar_fechamento_vs_data_emissao(data_emissao_fmt: str) -> None:
    """Valida se a data de emissão está dentro do período fiscal aberto no Sistema Fiscal."""
    data_fechamento = _obter_data_fechamento_sistema_fiscal()
    if not data_fechamento:
        return
    try:
        dt_emissao = datetime.strptime(str(data_emissao_fmt).strip(), "%d/%m/%Y")
        dt_fechamento = datetime.strptime(str(data_fechamento).strip(), "%d/%m/%Y")
        if dt_emissao > dt_fechamento:
            raise ErroNBS(
                "Período fiscal fechado no NBS. "
                f"Fechamento atual: {data_fechamento} | Emissão da solicitação: {data_emissao_fmt}. "
                "Solicitar abertura/atualização do fechamento no NBS antes de lançar Entradas."
            )
        bot.logger.informar(
            f"Fechamento fiscal identificado: {data_fechamento} (compatível com emissão {data_emissao_fmt})"
        )
    except ValueError:
        bot.logger.informar(f"Data de fechamento detectada, mas não foi possível validar: {data_fechamento}")


def obter_codigo_uab_contabil(tipo_pagamento_holmes: str) -> str | None:
    """
    TE07: Consulta planilha TIPOS_DE_PAGAMENTO.
    Coluna A = Tipo de Pagamento Holmes → Coluna D = Código UAB (CODIGO_UAB_CONTABIL).
    Retorna None se não encontrado (TE07: não lançar, Business Exception).
    """
    try:
        from pathlib import Path
        caminho = bot.configfile.obter_opcao_ou("nbs", "planilha_tipos_pagamento")
        if caminho and not Path(caminho).is_absolute():
            raiz = Path(__file__).resolve().parent.parent.parent
            candidato = (raiz / caminho).resolve()
            if candidato.exists():
                caminho = str(candidato)
        if not caminho or not bot.windows.afirmar_arquivo(caminho):
            raiz = Path(__file__).resolve().parent.parent.parent
            for nome in ("TIPOS_DE_PAGAMENTO.xlsx", "TIPOS_DE_PAGAMENTO.xls"):
                fallback = raiz / nome
                if fallback.exists():
                    caminho = str(fallback)
                    break
        if not caminho or not (Path(caminho).exists() or bot.windows.afirmar_arquivo(caminho)):
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
    campos: CamposNBS,
    dados_processo: dict,
    parar_antes_confirmar: bool = False,
    parar_apos_observacoes: bool = False,
    codigo_uab_contabil: str | None = None,
) -> RetornoStatus:
    """
    Fluxo NBS: Admin → NBS Fiscal → Entrada Só Diversa (Passos 1-24 conforme Detalhamento_RPA).
    
    Processa UMA ÚNICA ENTRADA de Solicitação de Pagamento.
    Passo 25 (Repetir para todos lançamentos): Responsabilidade do chamador (implementar loop externo).
    
    campos: CamposNBS com cpf, cod_empresa, cod_filial, processo, iniciar_em,
            despesa_pagamento, protocolo, centro_custo, data_vencimento, tipo_pagamento.
    parar_antes_confirmar: Se True, para antes de clicar em Confirmar (nao cria entrada).
    parar_apos_observacoes: Se True, para após Total Nota e Observações (para testes parciais).
    """
    # Compatibilidade: aceitar lista posicional (legado) ou CamposNBS
    if isinstance(campos, list):
        campos = CamposNBS(*campos)

    cpf = campos.cpf
    cod_empresa = campos.cod_empresa
    cod_filial = campos.cod_filial
    processo = campos.processo
    iniciar_em = campos.iniciar_em
    despesa_pagamento = campos.despesa_pagamento
    protocolo = campos.protocolo
    centro_custo = campos.centro_custo
    data_vencimento = campos.data_vencimento
    tipo_pagamento = campos.tipo_pagamento

    # TE07: Tipo de Pagamento não encontrado na planilha
    codigo_uab = codigo_uab_contabil or obter_codigo_uab_contabil(tipo_pagamento)
    if not codigo_uab:
        return RetornoStatus(
            False,
            f"Tipo de Pagamento '{tipo_pagamento}' não parametrizado na planilha TIPOS_DE_PAGAMENTO. "
            "Não lançar no NBS. Registrar Business Exception.",
        )

    data_emissao = _normalizar_data_ddmmyyyy(iniciar_em)
    data_vencimento_fmt = _normalizar_data_ddmmyyyy(data_vencimento)
    if not data_emissao:
        return RetornoStatus(False, f"Data de Emissão/Entrada inválida: {iniciar_em}")
    if not data_vencimento_fmt:
        return RetornoStatus(False, f"Data de Vencimento inválida: {data_vencimento}")

    data_emissao_digits = re.sub(r"\D", "", data_emissao)[:8]
    data_vencimento_digits = re.sub(r"\D", "", data_vencimento_fmt)[:8]
    
    # Validar que as datas possuem exatamente 8 dígitos (DDMMYYYY)
    if len(data_emissao_digits) != 8:
        return RetornoStatus(False, f"Data de Emissão/Entrada inválida para NBS: {data_emissao}")
    if len(data_vencimento_digits) != 8:
        return RetornoStatus(False, f"Data de Vencimento inválida para NBS: {data_vencimento_fmt}")
    
    # Log diagnóstico: formatos exatos sendo usados
    bot.logger.informar(
        f"Datas normalizadas - Emissão: {data_emissao} → {data_emissao_digits} | "
        f"Vencimento: {data_vencimento_fmt} → {data_vencimento_digits}"
    )

    # Data de Corte: validar antes de abrir o NBS (evita trabalho desnecessário)
    data_corte = bot.configfile.obter_opcao_ou("nbs", "data_corte")
    bloquear_fora_corte = str(bot.configfile.obter_opcao_ou("nbs", "data_corte_bloquear") or "").lower() in ("1", "true", "sim", "s")
    if data_corte and bloquear_fora_corte:
        try:
            dt_corte = datetime.strptime(str(data_corte).strip(), "%d/%m/%Y")
            dt_venc = datetime.strptime(str(data_vencimento_fmt).strip(), "%d/%m/%Y")
            if dt_venc > dt_corte:
                msg = f"Data de Vencimento ({data_vencimento_fmt}) fora da Data de Corte ({data_corte})."
                bot.logger.alertar(msg)
                return RetornoStatus(False, msg)
        except (ValueError, TypeError):
            pass

    try:
        bot.logger.informar(f"Iniciando entrada Solicitação Pagamento - processo {processo}, protocolo {protocolo}")
        _aguardar_contexto_nbs("inicio do processamento NBS", timeout_s=15.0, preferir_entradas=False)
        _dispensar_popup_data_hora_invalida()
        titulo_nbs = bot.configfile.obter_opcao_ou("nbs", "titulo_sistema") or "NBS"
        titulos = bot.windows.Janela.titulos_janelas()
        lista_t = list(titulos) if isinstance(titulos, (list, tuple, set)) else ([titulos] if titulos else [])
        for t in lista_t:
            if t and (titulo_nbs in str(t) or "NBS" in str(t).upper()):
                bot.windows.Janela(t).focar()
                break
        else:
            j = _obter_janela_nbs_segura()
            if j:
                j.focar()
        _dispensar_popup_data_hora_invalida()

        # Aguardar Empresa/Filial, Sistema Fiscal, Notas Fiscais ou Entradas (fluxo NBS Fiscal)
        def _janela_empresa_filial():
            t = bot.windows.Janela.titulos_janelas()
            s = str(t)
            return "Empresa" in s or "Exercício" in s or "Sistema Fiscal" in s or "Notas Fiscais" in s or "Entradas" in s or "Entrada" in s

        bot.util.aguardar_condicao(_janela_empresa_filial, 15)
        titulos = bot.windows.Janela.titulos_janelas()
        titulos_str = str(titulos)
        lista_titulos_ef = list(titulos) if isinstance(titulos, (list, tuple, set)) else ([titulos] if titulos else [])

        # [1] Empresa/Filial/Exercício Contábil (doc: preencher conforme Holmes)
        # Exercício é automático no NBS. Aqui garantimos Empresa/Filial + Confirma.
        _confirmou_ef = _resolver_empresa_filial_com_retry(
            str(cod_empresa),
            str(cod_filial),
            tentativas=4,
        )
        if _confirmou_ef:
            _garantir_sem_popup_data_hora("apos confirmar Empresa/Filial", ciclos=10)
        else:
            # Se não encontrou dialog de Empresa/Filial, pode já estar em Sistema Fiscal/Entradas.
            # Se o dialog ainda estiver aberto e não confirmou, aborta com erro claro.
            if _obter_janela_empresa_filial() is not None:
                caminho_print = _capturar_print_erro("empresa_filial_sem_confirmar")
                msg_print = f" Print: {caminho_print}" if caminho_print else ""
                raise ErroNBS(f"Não foi possível clicar em Confirmar na janela Empresa/Filial.{msg_print}")

        if _janela_empresa_filial_aberta():
            _diagnostico_ui_snapshot("empresa_filial_ainda_aberta_apos_validacao_inicial", capturar_print=True)
            caminho_print = _capturar_print_erro("empresa_filial_ainda_aberta")
            msg_print = f" Print: {caminho_print}" if caminho_print else ""
            raise ErroNBS(
                "Janela Empresa/Filial ainda está aberta. Fluxo bloqueado para evitar cliques fora de contexto."
                f"{msg_print}"
            )
        
        # Resolver eventual modal tardia de Empresa/Filial antes de validar a navegação principal.
        if not _garantir_empresa_filial_resolvida_antes_entradas(str(cod_empresa), str(cod_filial), timeout_aparecer_s=6.0):
            _diagnostico_ui_snapshot("empresa_filial_tardia_nao_resolvida", capturar_print=True)
            caminho_print = _capturar_print_erro("empresa_filial_tardia")
            msg_print = f" Print: {caminho_print}" if caminho_print else ""
            raise ErroNBS(
                "Janela Empresa/Filial permaneceu aberta (detecção tardia) antes de acessar Entradas."
                f"{msg_print}"
            )

        # Aguardar Sistema Fiscal (ou estado já adiantado em Entradas) antes de seguir.
        sleep(1)
        if not _aguardar_sistema_fiscal_pronto(timeout_s=10.0):
            _diagnostico_ui_snapshot("sistema_fiscal_nao_ficou_pronto", capturar_print=True)
            raise ErroNBS("Sistema Fiscal/Entradas não ficou pronto a tempo antes do clique em Nota Fiscal Entradas.")
        _validar_fechamento_vs_data_emissao(data_emissao)

        # [3] Nota fiscal ENTRADAS - clicar no ícone apenas quando ainda não estivermos em Entradas.
        ja_em_entradas = _obter_janela_nota_fiscal_diversas() is not None
        if ja_em_entradas:
            bot.logger.informar("Janela Entradas já está aberta; pulando clique em Nota Fiscal Entradas.")

        if not ja_em_entradas:
            bot.logger.informar("Clicando em Nota fiscal entradas (modo robusto)")
            _dispensar_popup_listagem_notas(tentativas=3)
            _garantir_sem_popup_data_hora("antes de clicar Nota fiscal entradas", ciclos=8)
            clicou_nf_entradas = _clicar_nota_fiscal_entradas_robusto(
                tentativas=4,
                cod_empresa=str(cod_empresa),
                cod_filial=str(cod_filial),
            )
            if not clicou_nf_entradas:
                _diagnostico_ui_snapshot("falha_clicar_nota_fiscal_entradas", capturar_print=True)
                raise ErroNBS("Não foi possível localizar o ícone 'Nota fiscal entradas'. Verificar imagem nbs_nota_fiscal_entradas.png.")

        # Aguardar e focar explicitamente a tela NBS-Nota Fiscal Diversas antes de seguir.
        if not _aguardar_janela_nota_fiscal_diversas(timeout_s=10.0):
            _diagnostico_ui_snapshot("falha_abrir_nota_fiscal_diversas", capturar_print=True)
            raise ErroNBS(
                "Clique em 'Nota fiscal entradas' não abriu a janela 'NBS-Nota Fiscal Diversas'."
            )
        try:
            _janela_nf_diversas = _obter_janela_nota_fiscal_diversas()
            if _janela_nf_diversas:
                _janela_nf_diversas.set_focus()
                sleep(0.3)
        except Exception:
            pass

        if not _aguardar_entradas_pronta_para_incluir(timeout_s=8.0):
            _diagnostico_ui_snapshot("entradas_nao_ficou_pronta_para_incluir", capturar_print=True)
            raise ErroNBS("Janela Entradas abriu, mas não ficou pronta para clicar em Incluir Entrada.")

        titulos = bot.windows.Janela.titulos_janelas()
        _titulos_lc = str(titulos).lower()
        if not any(k in _titulos_lc for k in ("entradas", "entrada diversa", "nota fiscal diversas", "nbs-nota fiscal diversas")):
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
            _digitar_com_foco_nbs("n", "atalho admin n", preferir_entradas=False)
            _digitar_com_foco_nbs("b", "atalho admin b", preferir_entradas=False)
            _digitar_com_foco_nbs("s", "atalho admin s", preferir_entradas=False)
            _tecla_com_foco_nbs("enter", "atalho admin enter", preferir_entradas=False)
            sleep(2)

        # Incluir Entrada (Só Diversa) - segundo ícone na toolbar
        bot.logger.informar("Clicando em Incluir Entrada (Só Diversa)")
        clicou_incluir = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["incluir_entrada_diversa"], segundos=5)
        if not clicou_incluir:
            clicou_incluir = _clicar_incluir_entrada_diversa_por_coordenada()
        if not clicou_incluir:
            _tecla_com_foco_nbs("alt", "fallback incluir - ALT", preferir_entradas=True)
            _digitar_com_foco_nbs("e", "fallback incluir - E", preferir_entradas=True)
            _tecla_com_foco_nbs("down", "fallback incluir - DOWN 1", preferir_entradas=True)
            _tecla_com_foco_nbs("down", "fallback incluir - DOWN 2", preferir_entradas=True)
            _tecla_com_foco_nbs("enter", "fallback incluir - ENTER", preferir_entradas=True)

        # Repetir tratamento do popup de data/hora após clicar em Incluir (passo solicitado)
        _garantir_sem_popup_data_hora("apos clicar Incluir Entrada (So Diversa)", ciclos=8)
        _validar_fechamento_vs_data_emissao(data_emissao)

        def _janela_entrada_abriu() -> bool:
            ts = str(_titulos_janelas_seguro())
            return "Entrada" in ts or "Diversa" in ts

        abriu_entrada = bot.util.aguardar_condicao(_janela_entrada_abriu, 6)
        if not abriu_entrada:
            _dispensar_popup_data_hora_invalida(tentativas=5)
            bot.logger.alertar("Entrada (Só Diversa) não abriu na primeira tentativa. Repetindo clique de inclusão.")
            clicou_incluir = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["incluir_entrada_diversa"], segundos=5)
            if not clicou_incluir:
                clicou_incluir = _clicar_incluir_entrada_diversa_por_coordenada()
            if not clicou_incluir:
                _tecla_com_foco_nbs("alt", "retry incluir - ALT", preferir_entradas=True)
                _digitar_com_foco_nbs("e", "retry incluir - E", preferir_entradas=True)
                _tecla_com_foco_nbs("down", "retry incluir - DOWN 1", preferir_entradas=True)
                _tecla_com_foco_nbs("down", "retry incluir - DOWN 2", preferir_entradas=True)
                _tecla_com_foco_nbs("enter", "retry incluir - ENTER", preferir_entradas=True)
            abriu_entrada = bot.util.aguardar_condicao(_janela_entrada_abriu, 8)
            if not abriu_entrada:
                raise ErroNBS("Não abriu a janela 'Entrada (Só Diversa)' após clicar em Incluir.")

        # Aguardar abertura da janela de entrada (no lugar do sleep fixo)
        bot.util.aguardar_condicao(
            lambda: any(
                "Entrada" in str(t) or "Diversa" in str(t)
                for t in (_titulos_janelas_seguro() or [])
                if t
            ),
            15,
        )

        # Aguardar janela de entrada
        titulos = bot.windows.Janela.titulos_janelas()
        lista_titulos = list(titulos) if isinstance(titulos, (list, tuple, set)) else ([titulos] if titulos else [])
        janela_entrada = None
        titulo_janela_entrada = ""

        # Prioridade: janela específica de operação (mais estável para as abas Contabilização/Faturamento).
        try:
            _w_entrada_op = _obter_janela_entrada_diversas_operacao()
            if _w_entrada_op:
                titulo_janela_entrada = str(_w_entrada_op.window_text() or "").strip()
                if titulo_janela_entrada:
                    janela_entrada = bot.windows.Janela(titulo_janela_entrada)
        except Exception:
            janela_entrada = None

        if not janela_entrada:
            for t in lista_titulos:
                ts = str(t or "")
                if not ts:
                    continue
                tsl = ts.lower()
                if "entrada diversas" in tsl or "entrada diversa" in tsl or "nbs-nota fiscal diversas" in tsl:
                    janela_entrada = bot.windows.Janela(ts)
                    titulo_janela_entrada = ts
                    break
        if not janela_entrada:
            janela_entrada = _obter_janela_nbs_segura()
            if not janela_entrada:
                raise ErroNBS("Não conseguiu localizar janela de Entrada após múltiplas tentativas.")
            try:
                titulo_janela_entrada = str(janela_entrada.titulo() or "")
            except Exception:
                titulo_janela_entrada = ""

        # Fornecedor CNPJ/CPF (TE01: se não encontrado, pendenciar)
        bot.logger.informar("Preenchendo Fornecedor CPF")
        campos_edit = janela_entrada.elementos(class_name="TEdit", top_level_only=False)
        if campos_edit:
            bot.mouse.clicar_mouse(coordenada=Elemento(campos_edit[0]).coordenada)
        bot.teclado.digitar_teclado(cpf)
        bot.teclado.apertar_tecla("tab")  # Trazer registro (doc: Dar Tab para trazer o registro)
        sleep(2)

        # TE01: CNPJ/CPF não encontrado → Pendenciar tarefa Holmes. Motivo: CNPJ não encontrado.
        sleep(1.5)
        cnpj_nao_encontrado = False
        try:
            from pywinauto import Desktop
            for w in Desktop(backend="win32").windows():
                try:
                    t = (w.window_text() or "").lower()
                    # Popup de erro do NBS: título curto com "encontrado" ou "cadastrado"
                    if ("encontrado" in t or "cadastrado" in t) and len(t) < 60:
                        w.type_keys("{ENTER}")
                        cnpj_nao_encontrado = True
                        break
                except Exception:
                    continue
        except Exception:
            pass
        if not cnpj_nao_encontrado:
            img_erro = _encontrar_imagem(IMAGENS_FALLBACK["cnpj_nao_encontrado"], segundos=2)
            if img_erro:
                bot.teclado.apertar_tecla("enter")
                cnpj_nao_encontrado = True
        if cnpj_nao_encontrado:
            raise ErroNBS("CNPJ não encontrado.")

        # Número NF, Série E, Emissão, Entrada (doc: Processo→Nº NF, Série=E, Iniciar em→Emissão e Entrada)
        # Após o Tab do CPF (lookup), cursor cai direto em Nº NF — não há tabs extras necessários.
        bot.logger.informar(
            f"Preenchendo Número NF, Série, Emissão, Entrada | emissao={data_emissao} ({data_emissao_digits})"
        )
        _preencheu_por_label = _preencher_capa_por_labels(janela_entrada, processo, data_emissao_digits)
        if not _preencheu_por_label:
            bot.logger.alertar("Preenchimento por labels indisponível. Usando fallback por teclado/Tab.")
            bot.teclado.digitar_teclado(processo)
            bot.teclado.apertar_tecla("tab")
            bot.teclado.digitar_teclado("E")  # Série E
            bot.teclado.apertar_tecla("tab")
            bot.teclado.atalho_teclado(["ctrl", "a"])
            sleep(0.1)
            bot.teclado.apertar_tecla("delete")  # Usar DELETE ao invés de BACKSPACE
            sleep(0.1)
            bot.teclado.digitar_teclado(data_emissao_digits)  # máscara NBS: digitar apenas dígitos
            _dispensar_popup_data_hora_invalida(tentativas=2)  # Proteção após emissão
            bot.teclado.apertar_tecla("tab")
            bot.teclado.atalho_teclado(["ctrl", "a"])
            sleep(0.1)
            bot.teclado.apertar_tecla("delete")  # Usar DELETE ao invés de BACKSPACE
            sleep(0.1)
            bot.teclado.digitar_teclado(data_emissao_digits)  # Entrada = Iniciar em
            _dispensar_popup_data_hora_invalida(tentativas=2)  # Proteção após entrada
            sleep(0.3)

        _dispensar_popup_data_hora_invalida(tentativas=5)
        if _ha_popup_data_hora_invalida():
            raise ErroNBS("Popup de data/hora inválida permaneceu aberto após preencher a capa.")
        # NÃO usar Tab aqui — o checkbox é clicado diretamente via pywinauto,
        # Tabs extras deslocam foco para campos ISS e corrompem o preenchimento.
        
        # Sleep extra para VCL Delphi processar a validação de data após preencher
        sleep(0.5)
        _dispensar_popup_data_hora_invalida(tentativas=3)  # Double-check

        # Desmarcar "Quero esta nota no livro fiscal" — buscar TCheckBox diretamente via pywinauto
        _desmarcou_flag = False
        _cx2 = 0
        _cy2 = 0
        try:
            from pywinauto import Desktop
            import ctypes as _ct2
            for _wf2 in Desktop(backend="win32").windows():
                try:
                    for _cls_chk in ("TCheckBox", "TDBCheckBox", "TwwCheckBox"):
                        for _chk2 in _wf2.descendants(class_name=_cls_chk):
                            _txt2 = (_chk2.window_text() or "").lower()
                            if "livro" in _txt2:
                                _cr2 = _chk2.rectangle()
                                _cx2 = _cr2.left + (_cr2.width() // 2)
                                _cy2 = _cr2.top + (_cr2.height() // 2)
                                # BM_GETCHECK = 0x00F0: 0=desmarcado, 1=marcado
                                _chk_state2 = _ct2.windll.user32.SendMessageW(_chk2.handle, 0x00F0, 0, 0)
                                if _chk_state2 != 0:
                                    import pyautogui as _pag2
                                    _pag2.moveTo(_cx2, _cy2, duration=0.1)
                                    sleep(0.05)
                                    _pag2.click()
                                    sleep(0.2)
                                    bot.logger.informar(f"Flag 'livro fiscal' desmarcada em ({_cx2},{_cy2})")
                                else:
                                    bot.logger.informar("Flag 'livro fiscal' já desmarcada")
                                _desmarcou_flag = True
                                break
                        if _desmarcou_flag:
                            break
                    if _desmarcou_flag:
                        break
                except Exception:
                    continue
        except Exception as _ef3:
            bot.logger.informar(f"Flag livro fiscal: {_ef3}")
        if not _desmarcou_flag:
            bot.teclado.apertar_tecla("space")
            bot.logger.informar("Flag livro fiscal: fallback space")
        sleep(0.5)

        # Total Nota (Despesa Pagamento) e Observações (Protocolo)
        # Após o clique no checkbox via pyautogui o foco se perde — clicar diretamente no campo.
        # Estratégia: localizar TLabel com texto "Total Nota" e obter o campo de input adjacente.
        _campo_total_clicado = False
        try:
            from pywinauto import Desktop
            for _wt3 in Desktop(backend="win32").windows():
                try:
                    _tt3 = (_wt3.window_text() or "")
                    if "Entrada" not in _tt3 and "Diversa" not in _tt3:
                        continue
                    # Mapear todos os campos de input com posição
                    _flds3 = []
                    for _cls3 in ("TOvcNumericField", "TOvcCurrencyField", "TDBEdit", "TEdit"):
                        for _f3 in _wt3.descendants(class_name=_cls3):
                            try:
                                _r3 = _f3.rectangle()
                                _flds3.append((_r3.top, _r3.left, _r3, _f3))
                            except Exception:
                                pass
                    # Localizar TLabel "Total Nota" e achar o campo à sua direita/mesma linha
                    _total_fld = None
                    for _lbl in _wt3.descendants(class_name="TLabel"):
                        _ltxt = (_lbl.window_text() or "").lower()
                        if "total" in _ltxt and "nota" in _ltxt:
                            _lr = _lbl.rectangle()
                            bot.logger.informar(f"TLabel 'Total Nota' em ({_lr.left},{_lr.top})")
                            # Campo = input com Y próximo ao label e X à sua direita
                            _cands = [
                                (abs(_rt.top - _lr.top), _xl3, _fobj3)
                                for _yt3, _xl3, _rt, _fobj3 in _flds3
                                if abs(_rt.top - _lr.top) < 20 and _xl3 >= _lr.left
                            ]
                            if _cands:
                                _cands.sort(key=lambda x: (x[0], x[1]))
                                _total_fld = _cands[0][2]
                            break
                    # Fallback: campo numérico mais à direita na linha mais baixa da Capa
                    if _total_fld is None and _flds3:
                        _flds3.sort(key=lambda x: (-x[0], -x[1]))  # Y desc, X desc
                        _total_fld = _flds3[0][3]
                        bot.logger.informar("Total Nota: fallback campo mais inferior-direito")
                    if _total_fld:
                        _total_fld.click_input()
                        sleep(0.1)
                        _total_fld.type_keys("^a", with_spaces=False)
                        _total_fld.type_keys(despesa_pagamento, with_spaces=True)
                        sleep(0.2)
                        _campo_total_clicado = True
                        bot.logger.informar(f"Total Nota preenchido: {despesa_pagamento}")
                    break
                except Exception as _et3i:
                    bot.logger.informar(f"Total Nota (janela): {_et3i}")
                    continue
        except Exception as _et3:
            bot.logger.informar(f"Campo Total Nota: {_et3}")
        if not _campo_total_clicado:
            bot.teclado.digitar_teclado(despesa_pagamento)
            bot.logger.informar("Total Nota: fallback teclado")
        # Tab move foco de Total Nota para o campo Obs (sequência natural do formulário)
        bot.teclado.apertar_tecla("tab")
        bot.teclado.digitar_teclado(protocolo)
        sleep(1)

        if parar_apos_observacoes:
            bot.logger.informar("PARADO: Preenchido até Total Nota e Observações. Continuar depois.")
            return RetornoStatus(True, "Parado após Observações - aguardando continuação.")

        # Guia Contabilização → + (Incluir)
        bot.logger.informar("Preenchendo Contabilização")
        _dispensar_janela_notas_ativas_sair(tentativas=3)

        clicou_aba = False
        _pw_nbs = None
        _target_pc = None
        _pc_rect_global = None
        _content_top = None
        _tab_ch = []
        _tab_idx_f = 6  # índice padrão de Faturamento; atualizado ao escanear abas

        try:
            from pywinauto import Desktop
            import ctypes
            import pyautogui as _pag

            # 1) Localizar janela NBS
            try:
                _tit_diag = bot.windows.Janela.titulos_janelas()
                bot.logger.informar(f"Contabilizacao: titulos abertos antes da busca da janela NBS: {_tit_diag}")
            except Exception:
                pass

            # Primeiro fallback forte: reaproveitar a mesma janela de Entrada já validada neste fluxo.
            try:
                _titulo_ref = str(titulo_janela_entrada or "").strip().lower()
                if _titulo_ref and len(_titulo_ref) >= 3:
                    for _w in Desktop(backend="win32").windows():
                        try:
                            _titulo_w = (_w.window_text() or "").strip().lower()
                            # Evita casar com títulos vazios: string vazia é substring de qualquer texto.
                            if not _titulo_w:
                                continue
                            if (
                                _titulo_ref == _titulo_w
                                or _titulo_ref in _titulo_w
                                or _titulo_w in _titulo_ref
                            ):
                                _pw_nbs = _w
                                _pw_nbs.set_focus()
                                sleep(0.2)
                                bot.logger.informar(
                                    f"Contabilizacao: janela de Entrada reutilizada como âncora ('{_w.window_text()}')"
                                )
                                break
                        except Exception:
                            continue
            except Exception:
                pass

            # Primeiro, garantir foco em uma janela de Entrada/NBS, quando disponível.
            try:
                _titulos_abertos = bot.windows.Janela.titulos_janelas()
                _lista_tit = list(_titulos_abertos) if isinstance(_titulos_abertos, (list, tuple, set)) else ([_titulos_abertos] if _titulos_abertos else [])
                for _t in _lista_tit:
                    _ts = str(_t or "")
                    _tl = _ts.lower()
                    if not _ts:
                        continue
                    if "notas fiscais" in _tl and any(k in _tl for k in ("ativas", "denegadas", "canceladas", "rejeitadas")):
                        continue
                    if any(k in _tl for k in ("entrada diversas", "entrada diversa", "nbs-nota fiscal diversas")):
                        try:
                            bot.windows.Janela(_ts).focar()
                            sleep(0.2)
                            break
                        except Exception:
                            continue
            except Exception:
                pass

            if not _pw_nbs:
                # Prioridade: janela de operação da Entrada Diversas (mais aderente ao fluxo esperado).
                try:
                    _w_entrada_op = _obter_janela_entrada_diversas_operacao()
                    if _w_entrada_op:
                        _pw_nbs = _w_entrada_op
                        _pw_nbs.set_focus()
                        sleep(0.25)
                        bot.logger.informar(
                            f"Contabilizacao: janela de Operação priorizada ('{_pw_nbs.window_text()}')"
                        )
                except Exception:
                    pass

            if not _pw_nbs:
                for _w in Desktop(backend="win32").windows():
                    try:
                        _titulo_w = (_w.window_text() or "").strip().lower()
                        if "notas fiscais" in _titulo_w and any(k in _titulo_w for k in ("ativas", "denegadas", "canceladas", "rejeitadas")):
                            continue
                        _tem_tab_com_contabil = any(
                            "contabil" in (s.window_text() or "").lower()
                            for s in _w.descendants(class_name="TTabSheet")
                        )
                        _tem_pagecontrol = bool(_w.descendants(class_name="TPageControl"))
                        _titulo_compativel = any(
                            k in _titulo_w
                            for k in ("entrada diversas", "entrada diversa", "nota fiscal diversa", "nota fiscal diversas", "nbs-nota fiscal diversas")
                        )
                        if _tem_tab_com_contabil or (_titulo_compativel and _tem_pagecontrol) or ("nota fiscal diversa" in _titulo_w) or ("nota fiscal diversas" in _titulo_w):
                            _pw_nbs = _w
                            _pw_nbs.set_focus()
                            sleep(0.3)
                            break
                    except Exception:
                        continue

            # Fallback: usar janela de foreground atual quando a enumeração não encontrar.
            if not _pw_nbs:
                try:
                    _hwnd_fg = ctypes.windll.user32.GetForegroundWindow()
                    if _hwnd_fg:
                        _w_fg = Desktop(backend="win32").window(handle=_hwnd_fg)
                        if _w_fg and _w_fg.descendants(class_name="TPageControl"):
                            _pw_nbs = _w_fg
                            _pw_nbs.set_focus()
                            sleep(0.2)
                            bot.logger.informar("Contabilizacao: usando janela de foreground como fallback")
                except Exception:
                    pass

            if _pw_nbs:
                # 2) Encontrar TPageControl pai da aba Contabilização.
                # Coleta TODOS os TPageControl que contêm "contabiliz" e ordena por
                # contagem de filhos ASC: o com MENOS filhos é o externo (barra visível).
                # O interno (25 abas) seria encontrado primeiro em busca depth-first, mas
                # TCM_GETITEMRECT nele retorna 0 pois é cross-process sem buffer remoto.
                _all_pc_cands = []
                for _pc_scan in _pw_nbs.descendants(class_name="TPageControl"):
                    try:
                        _ch_scan = [c for c in _pc_scan.children()
                                    if (c.class_name() or "") == "TTabSheet"]
                        if any("contabil" in (s.window_text() or "").lower() for s in _ch_scan):
                            _all_pc_cands.append((len(_ch_scan), _pc_scan, _ch_scan))
                        elif len(_ch_scan) >= 2:
                            # Fallback: há abas, mas os captions podem não ser expostos pelo wrapper.
                            _all_pc_cands.append((len(_ch_scan), _pc_scan, _ch_scan))
                    except Exception:
                        continue
                # Ordenar: menor count = TPageControl externo (barra de abas visível)
                _all_pc_cands.sort(key=lambda x: x[0])
                bot.logger.informar(
                    f"TPageControls com Contabilização: "
                    f"{[(n, [s.window_text() for s in ch]) for n, _, ch in _all_pc_cands]}"
                )
                if _all_pc_cands:
                    _, _target_pc, _tab_ch = _all_pc_cands[0]
                    _pchWnd = _target_pc.handle
                    _tab_idx_c = 0
                    _achou_tab_contab = False
                    for _i, _s in enumerate(_tab_ch):
                        _sn = (_s.window_text() or "").lower()
                        if "contabil" in _sn:
                            _tab_idx_c = _i
                            _achou_tab_contab = True
                        if "faturamento" in _sn:
                            _tab_idx_f = _i
                    if not _achou_tab_contab:
                        bot.logger.informar("Contabilizacao: caption da aba não exposto; usando índice padrão 0")

                if _target_pc:
                    _pc_rect_global = _target_pc.rectangle()
                    _pchWnd = _target_pc.handle
                    try:
                        _content_top = _tab_ch[0].rectangle().top
                    except Exception:
                        _content_top = _pc_rect_global.top + 24

                    _tab_names = [(_s.window_text() or "").strip() for _s in _tab_ch]
                    bot.logger.informar(
                        f"TPageControl externo: {_tab_names} | idx Contab={_tab_idx_c} "
                        f"| rect=({_pc_rect_global.left},{_pc_rect_global.top},"
                        f"{_pc_rect_global.right},{_pc_rect_global.bottom})"
                        f"| content_top={_content_top}"
                    )

                    # 3) Ativar aba Contabilização via TCM_SETCURSEL (0x130C).
                    # A mensagem simplesmente define a aba ativa pelo índice — sem ponteiro,
                    # sem buffer cross-process, sem coordenadas. É a forma mais direta e robusta.
                    # Após SetCurSel, disparar WM_NOTIFY/TCN_SELCHANGE para o pai, forçando
                    # que o NBS processe a troca de aba visualmente.
                    def _validador_aba_contabilizacao() -> bool:
                        try:
                            import ctypes as _ct_v
                            _cur = _ct_v.windll.user32.SendMessageW(_pchWnd, 0x130B, 0, 0)
                            if _cur == _tab_idx_c:
                                return True
                        except Exception:
                            pass
                        return _encontrar_imagem(
                            IMAGENS_FALLBACK["botao_incluir"],
                            confianca=0.8,
                            segundos=1,
                        ) is not None

                    _acc_ok = False
                    try:
                        import ctypes as _ct_tab
                        import pyautogui as _pag

                        _TCM_SETCURSEL = 0x130C
                        _prev = _ct_tab.windll.user32.SendMessageW(
                            _pchWnd, _TCM_SETCURSEL, _tab_idx_c, 0
                        )
                        bot.logger.informar(
                            f"TCM_SETCURSEL Contab [{_tab_idx_c}]: prev={_prev}"
                        )
                        # TCN_SELCHANGE = -551 (0xFFFFFDD9); WM_NOTIFY = 0x004E
                        # Enviar notificação ao pai para que o NBS atualize o conteúdo da aba
                        _parent_hwnd = _ct_tab.windll.user32.GetParent(_pchWnd)

                        class _NMHDR(_ct_tab.Structure):
                            _fields_ = [
                                ("hwndFrom", _ct_tab.c_void_p),
                                ("idFrom", _ct_tab.c_ulong),
                                ("code", _ct_tab.c_int),
                            ]

                        _nm = _NMHDR()
                        _nm.hwndFrom = _pchWnd
                        _nm.idFrom = _ct_tab.windll.user32.GetDlgCtrlID(_pchWnd)
                        _nm.code = -551  # TCN_SELCHANGE
                        _ct_tab.windll.user32.SendMessageW(
                            _parent_hwnd, 0x004E, _nm.idFrom, _ct_tab.byref(_nm)
                        )
                        sleep(0.5)
                        # Verificar se a aba foi realmente ativada
                        _cur = _ct_tab.windll.user32.SendMessageW(
                            _pchWnd, 0x130B, 0, 0  # TCM_GETCURSEL
                        )
                        bot.logger.informar(f"TCM_GETCURSEL após SetCurSel: {_cur}")
                        if _cur == _tab_idx_c:
                            _acc_ok = True
                            clicou_aba = True
                            bot.logger.informar(
                                f"Aba Contabilização [{_tab_idx_c}] ativada via TCM_SETCURSEL"
                            )
                        else:
                            # Fallback: clicar na aba usando PostMessage WM_LBUTTONDOWN
                            # na posição central do TTabSheet (header sempre visível no multirow)
                            bot.logger.informar(
                                f"TCM_SETCURSEL não ativou (cur={_cur}), tentando click direto"
                            )
                    except Exception as _etcm:
                        bot.logger.informar(f"TCM_SETCURSEL Contab: {_etcm}")

                    if not _acc_ok:
                        bot.logger.informar("TCM falhou, tentando imagem Contabilização com validação")
                        clicou_aba = _ativar_aba_por_imagem_robusta(
                            "aba_contabilizacao",
                            "Contabilização",
                            tentativas=4,
                            confianca=0.85,
                            segundos_busca=3,
                            validador=_validador_aba_contabilizacao,
                        )
                else:
                    bot.logger.informar("TPageControl com aba Contabilização não localizado")
            else:
                bot.logger.informar("Janela NBS não localizada")

        except Exception as _e:
            bot.logger.informar(f"Click aba Contabilização: {_e}")

        if not clicou_aba:
            clicou_aba = _ativar_aba_por_imagem_robusta(
                "aba_contabilizacao",
                "Contabilização",
                tentativas=3,
                confianca=0.85,
                segundos_busca=3,
            )
        if not clicou_aba:
            raise ErroNBS("Não foi possível localizar aba Contabilização")
        sleep(1.5)  # Aguardar VCL Delphi renderizar conteúdo da aba

        # 4) Botão "+" — após ativar a aba, buscar por posição visual
        # O "+" fica na seção "Contabilização Padrão", parte inferior da aba.
        # Estratégia: varrer TSpeedButton/TBitBtn pela posição Y mais baixa.
        clicou_incluir = False
        if _pw_nbs:
            try:
                import pyautogui as _pag
                _btn_candidatos = []
                for _cls in ("TSpeedButton", "TBitBtn", "TButton"):
                    for _btn in _pw_nbs.descendants(class_name=_cls):
                        try:
                            _br = _btn.rectangle()
                            _txt = (_btn.window_text() or "").strip()
                            # Filtrar: apenas botões dentro da área do TPageControl (ou abaixo)
                            if _pc_rect_global and _br.top >= _pc_rect_global.top:
                                _btn_candidatos.append((_br.top, _br.left, _br, _txt, _cls))
                        except Exception:
                            continue
                # Ordenar por posição Y (mais próximo da base = o "+" e os botões de contabilização)
                _btn_candidatos.sort(key=lambda x: (x[0], x[1]))
                _btns_diag = [f"{c[4]}@({c[2].left},{c[2].top})='{c[3]}'" for c in _btn_candidatos[:20]]
                bot.logger.informar(f"Botões na área Contab: {_btns_diag}")
                # Clicar no primeiro que tiver "+" OU que esteja na linha inferior (botões de ação)
                for _btop, _bleft, _br2, _btxt, _bcls in _btn_candidatos:
                    if "+" in _btxt:
                        _cx = _br2.left + (_br2.width() // 2)
                        _cy = _br2.top + (_br2.height() // 2)
                        _pag.moveTo(_cx, _cy, duration=0.15)
                        sleep(0.1)
                        _pag.click()
                        clicou_incluir = True
                        bot.logger.informar(f"Clicou '+' em ({_cx},{_cy})")
                        break
            except Exception as _e4:
                bot.logger.informar(f"Busca botão +: {_e4}")

        if not clicou_incluir:
            clicou_incluir = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["botao_incluir"])

        # Fallback posicional: "+" está ~55% horizontal, ~85% vertical da área de conteúdo
        if not clicou_incluir and _pc_rect_global is not None and _content_top is not None:
            try:
                import pyautogui as _pag
                _btn_x = _pc_rect_global.left + int(_pc_rect_global.width() * 0.55)
                _btn_y = _content_top + int((_pc_rect_global.bottom - _content_top) * 0.87)
                bot.logger.informar(f"Fallback posição '+': ({_btn_x},{_btn_y})")
                _pag.moveTo(_btn_x, _btn_y, duration=0.15)
                sleep(0.1)
                _pag.click()
                clicou_incluir = True
            except Exception as _e5:
                bot.logger.informar(f"Fallback posição +: {_e5}")

        if not clicou_incluir:
            raise ErroNBS("Não foi possível localizar botão Incluir contabilização")

        # Janela "Incluir Conta de Contabilização" (doc passo 13-14):
        # - Campo "Conta Contábil" = codigo_uab  (ex: 6123034)
        # - Campo "Centro de Custo" = codigo_uab  (doc passo 15: mesmo código UAB)
        # - Clicar "Confirmar"
        bot.util.aguardar_condicao(
            lambda: any(
                k in str(t)
                for t in (bot.windows.Janela.titulos_janelas() or [])
                for k in ("Incluir Conta", "Contabiliza", "Conta Cont", "Contábil", "Contabil",
                          "Atenção", "Atencao", "Bloqueio", "Aviso", "Alerta")
                if t
            ),
            8,
        )
        sleep(0.3)
        # Dispensar popups de aviso/bloqueio que o NBS abre antes do popup principal
        try:
            from pywinauto import Desktop as _DeskDismiss
            for _attempt_d in range(5):
                _dismissed_d = False
                for _wpop_d in _DeskDismiss(backend="win32").windows():
                    try:
                        _wpopt_d = (_wpop_d.window_text() or "")
                        if any(k in _wpopt_d.lower() for k in ("atenção", "atencao", "aviso", "bloqueio", "alerta")):
                            bot.logger.informar(f"Dismissing popup: '{_wpopt_d}'")
                            _btn_ok_d = None
                            for _cls_ok_d in ("TBitBtn", "TButton"):
                                for _bpop_d in _wpop_d.descendants(class_name=_cls_ok_d):
                                    if any(k in (_bpop_d.window_text() or "").lower() for k in ("ok", "confirm", "fechar", "sim", "yes")):
                                        _btn_ok_d = _bpop_d
                                        break
                                if _btn_ok_d:
                                    break
                            if _btn_ok_d:
                                _btn_ok_d.click_input()
                            else:
                                _wpop_d.type_keys("{ENTER}")
                            sleep(0.4)
                            _dismissed_d = True
                            break
                    except Exception:
                        continue
                if not _dismissed_d:
                    break
        except Exception:
            pass
        sleep(0.5)
        # Log diagnóstico: listar TODOS os títulos de janelas abertas após clicar '+'
        try:
            from pywinauto import Desktop as _DeskDiag
            _titulos_diag = [(_wd.window_text() or "") for _wd in _DeskDiag(backend="win32").windows() if (_wd.window_text() or "").strip()]
            bot.logger.informar(f"Janelas abertas após clicar '+': {_titulos_diag}")
        except Exception:
            pass
        # Localizar a janela popup via pywinauto (título pode variar)
        _pw_contab = None
        try:
            from pywinauto import Desktop
            # Aguardar popup Incluir Conta appear (pode demorar após fechar popups de aviso)
            bot.util.aguardar_condicao(
                lambda: any(
                    k in str(t)
                    for t in (bot.windows.Janela.titulos_janelas() or [])
                    for k in ("Incluir Conta", "Contabiliza", "Conta Cont", "Contábil", "Contabil")
                    if t
                ),
                5,
            )
            for _wc in Desktop(backend="win32").windows():
                try:
                    _wct = (_wc.window_text() or "").strip()
                    if any(k in _wct for k in ("Incluir Conta", "Contabiliza", "Conta Cont")):
                        _pw_contab = _wc
                        _pw_contab.set_focus()
                        bot.logger.informar(f"Janela contabilização: '{_wct}'")
                        break
                except Exception:
                    continue
        except Exception:
            pass

        if _pw_contab:
            try:
                # Preencher Conta Contábil (primeiro TEdit) com codigo_uab
                _edits_c = [e for e in _pw_contab.descendants(class_name="TEdit")]
                if _edits_c:
                    _edits_c[0].click_input()
                    sleep(0.1)
                    _edits_c[0].type_keys("^a", with_spaces=False)
                    _edits_c[0].type_keys(codigo_uab, with_spaces=True)
                    sleep(0.3)
                bot.logger.informar(f"Conta Contábil preenchida: {codigo_uab}")
                # Tab para Centro de Custo (segundo campo editável) = mesmo codigo_uab (doc passo 15)
                import pyautogui as _pag
                _pag.hotkey("tab")
                sleep(0.3)
                if len(_edits_c) >= 2:
                    _edits_c[1].click_input()
                    sleep(0.1)
                    _edits_c[1].type_keys("^a", with_spaces=False)
                    _edits_c[1].type_keys(codigo_uab, with_spaces=True)
                else:
                    _pag.typewrite(codigo_uab, interval=0.05)
                sleep(0.3)
                bot.logger.informar(f"Centro de Custo preenchido: {codigo_uab}")
                # Clicar Confirmar
                _confirmou = False
                for _cls_b in ("TBitBtn", "TButton"):
                    for _bc in _pw_contab.descendants(class_name=_cls_b):
                        _bt = (_bc.window_text() or "").lower()
                        if "confirm" in _bt or "ok" in _bt:
                            _bc.click_input()
                            _confirmou = True
                            break
                    if _confirmou:
                        break
                if not _confirmou:
                    _pag.hotkey("enter")
                bot.logger.informar("Contabilização confirmada")
            except Exception as _ec:
                bot.logger.informar(f"Preencher janela contabilização: {_ec}")
                bot.teclado.apertar_tecla("enter")
        else:
            # Fallback teclado: Conta Contábil + tab + Centro de Custo + enter
            bot.logger.informar("Popup contabilização não encontrado — usando fallback teclado")
            bot.teclado.digitar_teclado(codigo_uab)
            bot.teclado.apertar_tecla("tab")
            bot.teclado.digitar_teclado(codigo_uab)
            bot.teclado.apertar_tecla("enter")
            sleep(1)
            bot.teclado.apertar_tecla("enter")

        # Aguardar retorno à janela principal após confirmar contabilização
        bot.util.aguardar_condicao(
            lambda: any(
                "Entrada" in str(t) or "Diversa" in str(t)
                for t in (bot.windows.Janela.titulos_janelas() or [])
                if t
            ),
            10,
        )
        sleep(0.5)

        # Guia Faturamento — TCM_SETCURSEL, mesmo mecanismo do bloco Contabilização
        bot.logger.informar("Preenchendo Faturamento")
        clicou_fat = False

        def _validador_aba_faturamento() -> bool:
            if _target_pc:
                try:
                    import ctypes as _ct_vf
                    _cur = _ct_vf.windll.user32.SendMessageW(_target_pc.handle, 0x130B, 0, 0)
                    if _cur == _tab_idx_f:
                        return True
                except Exception:
                    pass
            return _encontrar_imagem(
                IMAGENS_FALLBACK["botao_gerar"],
                confianca=0.8,
                segundos=1,
            ) is not None

        if _pw_nbs and _target_pc:
            try:
                import ctypes as _ctf
                _pchWndF = _target_pc.handle
                _prev_f = _ctf.windll.user32.SendMessageW(
                    _pchWndF, 0x130C, _tab_idx_f, 0  # TCM_SETCURSEL
                )
                bot.logger.informar(f"TCM_SETCURSEL Faturamento [{_tab_idx_f}]: prev={_prev_f}")
                _parent_hwnd_f = _ctf.windll.user32.GetParent(_pchWndF)

                class _NMHDR_F(_ctf.Structure):
                    _fields_ = [
                        ("hwndFrom", _ctf.c_void_p),
                        ("idFrom", _ctf.c_ulong),
                        ("code", _ctf.c_int),
                    ]

                _nm_f = _NMHDR_F()
                _nm_f.hwndFrom = _pchWndF
                _nm_f.idFrom = _ctf.windll.user32.GetDlgCtrlID(_pchWndF)
                _nm_f.code = -551  # TCN_SELCHANGE
                _ctf.windll.user32.SendMessageW(
                    _parent_hwnd_f, 0x004E, _nm_f.idFrom, _ctf.byref(_nm_f)
                )
                sleep(0.5)
                _cur_f = _ctf.windll.user32.SendMessageW(_pchWndF, 0x130B, 0, 0)
                bot.logger.informar(f"TCM_GETCURSEL após Faturamento: {_cur_f}")
                if _cur_f == _tab_idx_f:
                    clicou_fat = True
                    bot.logger.informar(
                        f"Aba Faturamento [{_tab_idx_f}] ativada via TCM_SETCURSEL"
                    )
                    # TCM muda o índice mas VCL Delphi só renderiza os controles com clique físico
                    _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["aba_faturamento"], segundos=4)
            except Exception as _ef:
                bot.logger.informar(f"TCM_SETCURSEL Faturamento: {_ef}")
        if not clicou_fat:
            clicou_fat = _ativar_aba_por_imagem_robusta(
                "aba_faturamento",
                "Faturamento",
                tentativas=4,
                confianca=0.85,
                segundos_busca=3,
                validador=_validador_aba_faturamento,
            )
        if not clicou_fat:
            raise ErroNBS("Não foi possível localizar aba Faturamento")
        sleep(2)  # Aguardar VCL Delphi renderizar conteúdo da aba
        # Total Parcelas = 1 e botão Gerar — usar _pw_nbs direto
        if _pw_nbs:
            try:
                _nums = _pw_nbs.descendants(class_name="TOvcNumericField")
                if _nums:
                    _nums[0].click_input()
                    _nums[0].type_keys("^a{BACKSPACE}1", with_spaces=False)
            except Exception:
                pass
        clicou_gerar = False
        if _pw_nbs:
            try:
                _btns_fat_diag = [
                    f"{_b.class_name()}@({_b.rectangle().left},{_b.rectangle().top})='{_b.window_text()}'"
                    for _b in _pw_nbs.descendants()
                    if (_b.class_name() or "") in ("TBitBtn", "TButton", "TSpeedButton")
                ]
                bot.logger.informar(f"Botões Faturamento (todos): {_btns_fat_diag[:25]}")
            except Exception as _ediag_g:
                bot.logger.informar(f"Diag botões Faturamento: {_ediag_g}")
        if _pw_nbs:
            try:
                for _cls in ("TBitBtn", "TButton", "TSpeedButton"):
                    for _bg in _pw_nbs.descendants(class_name=_cls):
                        if "gerar" in (_bg.window_text() or "").lower():
                            _bg.click_input()
                            clicou_gerar = True
                            break
                    if clicou_gerar:
                        break
            except Exception:
                pass
        if not clicou_gerar:
            clicou_gerar = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["botao_gerar"])
        if not clicou_gerar:
            raise ErroNBS("Não foi possível localizar botão Gerar")
        sleep(2)

        # Tipo Pagamento = Boleto, Natureza Despesa = Outras Despesas
        if _pw_nbs:
            try:
                _combos = []
                for _cls_combo in ("TwwDBLookupCombo", "TDBLookupComboBox", "TComboBox"):
                    _all_combos_f = _pw_nbs.descendants(class_name=_cls_combo)
                    # Filtrar apenas combos visíveis — os de outras abas (invisíveis) causam exceção
                    _combos = [c for c in _all_combos_f if c.is_visible()]
                    if _combos:
                        bot.logger.informar(f"Combos Faturamento ({_cls_combo}): {len(_all_combos_f)} total, {len(_combos)} visíveis")
                        break
                if not _combos:
                    bot.logger.informar("Combos Faturamento: nenhum combo visível encontrado")
                if len(_combos) >= 2:
                    # Log diagnóstico: exibir texto atual e posição de cada combo visível
                    # para confirmar qual é Tipo Pagamento (esperado: combo[0]) e qual é
                    # Natureza Despesa (esperado: combo[1]).
                    for _ic, _cc in enumerate(_combos):
                        try:
                            _cr = _cc.rectangle()
                            _cv = (_cc.window_text() or _cc.texts()[0] if _cc.texts() else "")
                            bot.logger.informar(
                                f"Combo[{_ic}] pos=({_cr.left},{_cr.top}) texto='{_cv}'"
                            )
                        except Exception:
                            pass
                    # Doc Passo 20: combo[0] = Tipo Pagamento → Boleto
                    #               combo[1] = Natureza Despesa → Outras Despesas
                    # (ordem topo→baixo conforme layout visual do NBS Faturamento)
                    _combos[0].click_input()
                    _combos[0].type_keys("^aBoleto{TAB}", with_spaces=True)
                    _combos[1].click_input()
                    _combos[1].type_keys("^aOutras Despesas{ENTER}", with_spaces=True)
                elif len(_combos) == 1:
                    _combos[0].click_input()
                    _combos[0].type_keys("^aBoleto{TAB}", with_spaces=True)
            except Exception as _ecb:
                bot.logger.informar(f"Combos Faturamento: {_ecb}")
        sleep(1)

        # Vencimento ← Data de Vencimento Holmes (máscara NBS: digitar apenas dígitos)
        # Campo TOvcDbPictureField com validação de data inválida pós-preenchimento
        _preencheu_vencimento = False
        if _pw_nbs:
            try:
                _flds_venc = _pw_nbs.descendants(class_name="TOvcDbPictureField")
                if _flds_venc:
                    # Em Faturamento, procura por label "Vencimento" e correlaciona com campo
                    # Fallback: assume que, após gerar parcelas, o primeiro TOvcDbPictureField é o vencimento
                    _campo_venc = None
                    try:
                        for _lbl in _pw_nbs.descendants(class_name="TLabel"):
                            _ltxt = (_lbl.window_text() or "").lower()
                            if "vencimento" in _ltxt or "vencimento boleto" in _ltxt:
                                _lr = _lbl.rectangle()
                                # Campo vencimento deve estar próximo (na mesma linha ou logo abaixo)
                                _cands_venc = [
                                    _f for _f in _flds_venc
                                    if abs(_f.rectangle().top - _lr.top) <= 25
                                ]
                                if _cands_venc:
                                    _campo_venc = _cands_venc[0]
                                break
                    except Exception:
                        pass
                    
                    # Fallback: usar primeiro TOvcDbPictureField visível se label não encontrado
                    if not _campo_venc:
                        _campo_venc = _flds_venc[0] if _flds_venc else None
                    
                    if _campo_venc:
                        # Limpeza agressiva: múltiplas tentativas de clear
                        _campo_venc.click_input()
                        sleep(0.15)
                        for _ in range(3):
                            _campo_venc.type_keys("^a{DELETE}", with_spaces=False)
                            sleep(0.05)
                        sleep(0.1)
                        # Preencher dígitos (máscara: DDMMYYYY)
                        _campo_venc.type_keys(data_vencimento_digits, with_spaces=False)
                        sleep(0.3)
                        bot.logger.informar(f"Vencimento preenchido: {data_vencimento_digits}")
                        _preencheu_vencimento = True
            except Exception as _evenc:
                bot.logger.informar(f"Vencimento preenchimento: {_evenc}")
        
        # Validar popup de data inválida após vencimento
        if _preencheu_vencimento:
            sleep(0.5)
            _dispensar_popup_data_hora_invalida(tentativas=5)
            if _ha_popup_data_hora_invalida():
                raise ErroNBS("Popup de data/hora inválida no campo Vencimento após preenchimento.")
            
            # Passo 21: Re-validar Data de Corte após preencher Vencimento no NBS
            data_corte = bot.configfile.obter_opcao_ou("nbs", "data_corte")
            bloquear_fora_corte = str(bot.configfile.obter_opcao_ou("nbs", "data_corte_bloquear") or "").lower() in ("1", "true", "sim", "s")
            if data_corte and bloquear_fora_corte:
                try:
                    dt_corte = datetime.strptime(str(data_corte).strip(), "%d/%m/%Y")
                    dt_venc = datetime.strptime(str(data_vencimento_fmt).strip(), "%d/%m/%Y")
                    if dt_venc > dt_corte:
                        msg = f"Passo 21: Data de Vencimento ({data_vencimento_fmt}) ainda fora da Data de Corte ({data_corte}) após preenchimento no NBS."
                        bot.logger.alertar(msg)
                        raise ErroNBS(msg)
                except ValueError:
                    bot.logger.informar("Re-validação de Data de Corte: parse falhou, continuando")
                except ErroNBS:
                    raise
                except Exception as _edc:
                    bot.logger.informar(f"Re-validação de Data de Corte: {_edc}")

        if parar_antes_confirmar:
            bot.logger.informar("PARADO: Formulario preenchido. Nao clicou em Confirmar (nao cria entrada).")
            return RetornoStatus(True, "Teste: parado antes de Confirmar.")

        # Confirmar (botão na barra inferior da janela principal)
        clicou_confirmar = False
        if _pw_nbs:
            try:
                for _cls in ("TBitBtn", "TButton"):
                    for _bc2 in _pw_nbs.descendants(class_name=_cls):
                        if "confirmar" in (_bc2.window_text() or "").lower():
                            _bc2.click_input()
                            clicou_confirmar = True
                            break
                    if clicou_confirmar:
                        break
            except Exception:
                pass
        if not clicou_confirmar:
            clicou_confirmar = _clicar_por_imagem_se_existir(IMAGENS_FALLBACK["botao_confirmar"])
        if not clicou_confirmar:
            raise ErroNBS("Não foi possível localizar botão Confirmar")

        # Aguardar pop-up de confirmação ou tela de Ficha de Controle
        bot.util.aguardar_condicao(
            lambda: any(
                k in str(bot.windows.Janela.titulos_janelas())
                for k in ("Aviso", "Informação", "Ficha", "Controle", "Número")
            ),
            10,
        )

        # Pop-up Aviso NF-e: capturar número de controle e fechar
        num_controle = None
        titulos_popup = ["Aviso", "Informação", "Número de Controle", "Nota Fiscal"]
        for titulo_pop in titulos_popup:
            if titulo_pop in str(bot.windows.Janela.titulos_janelas()):
                try:
                    janela_popup = bot.windows.Janela(titulo_pop)
                    for elem in janela_popup.elementos(top_level_only=False):
                        txt = (elem.window_text() or "").strip()
                        if txt and any(c.isdigit() for c in txt) and len(txt) < 50:
                            num_controle = txt
                            break
                except Exception:
                    pass
                bot.teclado.apertar_tecla("enter")
                break
        if num_controle:
            bot.logger.informar(f"Número de controle capturado: {num_controle}")
        else:
            bot.logger.informar("Pop-up de número de controle não detectado ou número não capturado.")

        # Ficha de Controle de Pagamento → Cancelar (doc passo 24: clicar Cancelar, não fechar pela X)
        sleep(2)
        titulos_ficha = bot.windows.Janela.titulos_janelas()
        if "Ficha" in str(titulos_ficha) or "Controle" in str(titulos_ficha):
            for t in (list(titulos_ficha) if isinstance(titulos_ficha, (list, tuple, set)) else ([titulos_ficha] if titulos_ficha else [])):
                if isinstance(t, str) and ("Ficha" in t or "Controle" in t):
                    janela_ficha = bot.windows.Janela(t)
                    bot.logger.informar(f"Ficha de Controle encontrada: '{t}'")
                    _cancelou_ficha = False
                    try:
                        from pywinauto import Desktop as _DeskFicha
                        for _wf in _DeskFicha(backend="win32").windows():
                            try:
                                _wft = (_wf.window_text() or "")
                                if "Ficha" in _wft or "Controle" in _wft:
                                    for _cls_f in ("TBitBtn", "TButton"):
                                        for _bf in _wf.descendants(class_name=_cls_f):
                                            _bft = (_bf.window_text() or "").lower()
                                            if "cancelar" in _bft or "cancel" in _bft:
                                                _bf.click_input()
                                                _cancelou_ficha = True
                                                bot.logger.informar("Ficha de Controle: clicou Cancelar")
                                                break
                                        if _cancelou_ficha:
                                            break
                                    break
                            except Exception:
                                continue
                    except Exception:
                        pass
                    if not _cancelou_ficha:
                        # Fallback: fechar pela X (comportamento anterior)
                        bot.logger.informar("Ficha de Controle: botão Cancelar não encontrado, fechando pela X")
                        janela_ficha.fechar()
                    break

        msg_controle = f"num_controle={num_controle}" if num_controle else "num_controle=não capturado"
        bot.logger.informar(f"Processamento concluído. {msg_controle}")
        return RetornoStatus(True, msg_controle)

    except ErroNBS as e:
        print_erro = _capturar_print_erro("erro_nbs")
        if print_erro:
            bot.logger.alertar(f"Print de erro salvo em: {print_erro}")
        return RetornoStatus(False, str(e))
    except Exception as e:
        print_erro = _capturar_print_erro("erro_inesperado")
        if print_erro:
            bot.logger.alertar(f"Print de erro salvo em: {print_erro}")
        bot.logger.alertar(f"Erro no processamento: {e}")
        return RetornoStatus(False, str(e))
