from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XToolkit3
from ooodev.adapter.awt.font_mapping_use_partial import FontMappingUsePartial
from ooodev.adapter.awt.toolkit2_partial import Toolkit2Partial


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
else:
    UnoInterface = Any


class Toolkit3Partial(
    Toolkit2Partial,
    FontMappingUsePartial,
):
    """
    Partial Class for XToolkit2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XToolkit3, interface: UnoInterface | None = XToolkit3) -> None:
        """
        Constructor

        Args:
            component (XToolkit3): UNO Component that implements ``com.sun.star.awt.XToolkit3`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XToolkit2``.
        """
        Toolkit2Partial.__init__(self, component=component, interface=interface)
        FontMappingUsePartial.__init__(self, component=component, interface=interface)
