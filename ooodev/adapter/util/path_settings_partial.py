from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.util import XPathSettings

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.util.path_settings_properties_partial import PathSettingsPropertiesPartial


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PathSettingsPartial(PathSettingsPropertiesPartial):
    """
    Partial Class XPathSettings.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPathSettings, interface: UnoInterface | None = XPathSettings) -> None:
        """
        Constructor

        Args:
            component (XPathSettings): UNO Component that implements ``com.sun.star.util.XPathSettings`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPathSettings``.
        """

        def validate(component: Any, interface: Any) -> None:
            if interface is None:
                return
            if not mLo.Lo.is_uno_interfaces(component, interface):
                raise mEx.MissingInterfaceError(interface)

        validate(component, interface)
        PathSettingsPropertiesPartial.__init__(self, component)


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
    builder.auto_add_interface("com.sun.star.util.XPathSettings", False)
    return builder
