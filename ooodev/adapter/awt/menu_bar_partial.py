from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XMenuBar
from ooodev.adapter.awt.menu_partial import MenuPartial


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
else:
    UnoInterface = Any


class MenuBarPartial(MenuPartial):
    """
    Partial Class for XMenuBar.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XMenuBar, interface: UnoInterface | None = XMenuBar) -> None:
        """
        Constructor

        Args:
            component (XMenuBar): UNO Component that implements ``com.sun.star.awt.XMenuBar`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMenuBar``.
        """
        MenuPartial.__init__(self, component=component, interface=interface)
