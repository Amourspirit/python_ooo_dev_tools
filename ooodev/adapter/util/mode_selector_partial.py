from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno
from com.sun.star.util import XModeSelector
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ModeSelectorPartial:
    """
    Partial Class XModeSelector.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XModeSelector, interface: UnoInterface | None = XModeSelector) -> None:
        """
        Constructor

        Args:
            component (XModeSelector): UNO Component that implements ``com.sun.star.util.XModeSelector`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XModeSelector``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XModeSelector
    def get_mode(self) -> str:
        """
        Gets the mode.
        """
        return self.__component.getMode()

    def get_supported_modes(self) -> Tuple[str, ...]:
        """
        Gets the supported modes.
        """
        return self.__component.getSupportedModes()

    def set_mode(self, mode: str) -> None:
        """
        Sets a new mode for the implementing object.

        Raises:
            com.sun.star.lang.NoSupportException: ``NoSupportException``
        """
        self.__component.setMode(mode)

    def supports_mode(self, mode: str) -> bool:
        """
        Gets whether a mode is supported or not.
        """
        return self.__component.supportsMode(mode)

    # endregion XModeSelector
