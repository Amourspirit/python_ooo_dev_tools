from __future__ import annotations
import unohelper

from typing import TYPE_CHECKING
from com.sun.star.awt import XTopWindowListener
if TYPE_CHECKING:
    from com.sun.star.lang import EventObject

class XTopWindowAdapter(unohelper.Base, XTopWindowListener):
    def windowOpened(event: EventObject) -> None:
        pass
    def windowActivated(event: EventObject) -> None:
        pass
    def windowDeactivated(event: EventObject) -> None:
        pass
    def windowMinimized(event: EventObject) -> None:
        pass
    def windowNormalized(event: EventObject) -> None:
        pass
    def windowClosing(event: EventObject) -> None:
        pass
    def windowClosed(event: EventObject) -> None:
        pass
    def disposing(event: EventObject) -> None:
        pass
