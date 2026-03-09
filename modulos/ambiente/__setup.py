import bot
import ctypes
import psutil
from typing import Literal
import getpass

def capslock_switcher( switch_to: Literal['ON', 'OFF'] ) -> None:
    status = ctypes.WinDLL("User32.dll").GetKeyState(0x14)
    if (status and switch_to == "OFF") or (not status and switch_to == "ON"):
        return bot.teclado.apertar_tecla('caps_lock')

def verificar_processos() -> list:
    """Verifica se existem processos do NBS em execução"""
    bot.logger.informar(f"Verificando processos em execução...")
    PROCESSOS_NBS = bot.configfile.obter_opcoes("nbs", [ str('processos').lower() ])[0].split(", ")
    usuario_atual = getpass.getuser()
    todos_processos = psutil.process_iter( ['pid', 'name', 'username'] )
    em_execucao = [
        processo for processo in todos_processos 
        if str(usuario_atual) in str(processo.info['username'])
        and any(p in processo.info['name'] for p in PROCESSOS_NBS)]

    return em_execucao

def encerrar_processos() -> bool:
    """Encerra os processos do NBS em execução"""
    try:
        #Verifica se existem processos em aberto
        em_execucao: list[psutil.Process] = verificar_processos()

        if em_execucao:
            bot.logger.informar(f"Encerrando { len( em_execucao ) } processo(s) em execução...")
            for processo in em_execucao:
                processo.terminate()
                bot.logger.informar(f"Processo(s) { [p.pid for p in em_execucao] } encerrado(s) com sucesso!")
            return True
        
        return True
    except Exception as erro:
        bot.logger.informar(f"Erro ao encerrar processos do NBS")
        return False

__all__ = [
    "encerrar_processos",
    "capslock_switcher"
]