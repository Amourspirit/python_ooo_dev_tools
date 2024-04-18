from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.lang import XLocalizable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.lang import Locale
    from ooodev.utils.type_var import UnoInterface


class LocalizablePartial:
    """
    Partial class for ``XLocalizable``.
    """

    def __init__(self, component: XLocalizable, interface: UnoInterface | None = XLocalizable) -> None:
        """
        Constructor

        Args:
            component (XLocalizable): UNO Component that implements ``com.sun.star.lang.XLocalizable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLocalizable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XLocalizable
    def get_locale(self) -> Locale:
        """
        returns the locale of the object.
        """
        return self.__component.getLocale()

    def setLocale(self, value: Locale) -> None:
        """
        Sets the locale to be used by this object.
        """
        self.__component.setLocale(value)

    # endregion XLocalizable


def get_builder(component: Any) -> Any:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.lang.XLocalizable", False)
    return builder
