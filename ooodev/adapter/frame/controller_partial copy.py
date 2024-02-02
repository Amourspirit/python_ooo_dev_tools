from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import theDesktop

from .desktop2_partial import Desktop2Partial

if TYPE_CHECKING:
    from com.sun.star.frame import XFrame
    from com.sun.star.frame import XModel
    from ooodev.utils.type_var import UnoInterface


class TheDesktopPartial(Desktop2Partial):
    """
    Partial class for theDesktop singleton class.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: theDesktop, interface: UnoInterface | None = theDesktop) -> None:
        """
        Constructor

        Args:
            component (theDesktop ): UNO Component that implements ``com.sun.star.frame.theDesktop`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``theDesktop``.
        """
        Desktop2Partial.__init__(self, component=component, interface=interface)
