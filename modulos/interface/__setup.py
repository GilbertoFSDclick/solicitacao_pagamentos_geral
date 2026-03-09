# interno
from bot.estruturas import Coordenada

# externo
from pywinauto.controls.hwndwrapper import HwndWrapper

class Elemento:
    """Elemento da interface do NBS"""

    elemento: HwndWrapper
    """Elemento do módulo `pywinauto`"""
    coordenada: Coordenada
    """Coordenada do `elemento`"""

    def __init__ (self, elemento: HwndWrapper) -> None:
        self.elemento = elemento
        box = self.elemento.rectangle()
        self.coordenada = Coordenada(box.left, box.top, box.width(), box.height())

    def descendentes (self, ativo=True) -> list[HwndWrapper]:
        """Filhos visíveis e `ativo` do `elemento`"""
        return [
            filho
            for filho in self.elemento.children()
            if filho.is_enabled() == ativo and filho.is_visible()
        ]

    def focar (self) -> None:
        """Focar no elemento"""
        self.elemento.set_focus()

__all__ = [
    "Elemento"
]