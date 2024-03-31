from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import contextlib
import uno

from com.sun.star.awt import XFontMappingUse

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from com.sun.star.awt import XFontMappingUseItem  # struct, would think this is an it is in fact a struct.
    from ooodev.utils.type_var import UnoInterface


class FontMappingUsePartial:
    """
    Partial class for XFontMappingUse.

    Since LibreOffice ``7.3``
    """

    def __init__(self, component: XFontMappingUse, interface: UnoInterface | None = XFontMappingUse) -> None:
        """
        Constructor

        Args:
            component (XFontMappingUse): UNO Component that implements ``com.sun.star.awt.XFontMappingUse`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFontMappingUse``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        if mInfo.Info.version_info >= (7, 3):
            validate(component, interface)
        self.__component = component

    # region XFontMappingUse
    def finish_tracking_font_mapping_use(self) -> Tuple[XFontMappingUseItem, ...]:
        """
        Stop tracking of how requested fonts are mapped to available fonts and return the mappings that took place since the call to ``start_tracking_font_mapping_use()``.

        Since LibreOffice ``7.3``
        """
        with contextlib.suppress(Exception):
            return self.__component.finishTrackingFontMappingUse()
        return ()

    def start_tracking_font_mapping_use(self) -> None:
        """
        Activate tracking of how requested fonts are mapped to available fonts.

        Since LibreOffice ``7.3``
        """
        with contextlib.suppress(Exception):
            self.__component.startTrackingFontMappingUse()

    # endregion XFontMappingUse
