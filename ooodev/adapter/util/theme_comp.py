from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.component_prop import ComponentProp

from ooodev.adapter.util.theme_partial import ThemePartial
from ooodev.adapter.util.theme_properties_partial import ThemePropertiesPartial

if TYPE_CHECKING:
    from com.sun.star.util import XTheme  # type: ignore


class ThemeComp(ComponentProp, ThemePartial, ThemePropertiesPartial):
    """
    Class for managing XTheme Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTheme | tuple) -> None:
        """
        Constructor

        Args:
            component (thePathSettings): UNO Component that implements ``com.sun.star.ui.thePathSettings`` service.
        """
        ComponentProp.__init__(self, component)
        ThemePartial.__init__(self, component=component, interface=None)
        ThemePropertiesPartial.__init__(self, component=component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # Avoid pre 7.6 versions of LibreOffice issues
        return ()

    # endregion Overrides

    # region Properties
    @property
    @override
    def component(self) -> XTheme:
        """thePathSettings Component"""
        # pylint: disable=no-member
        return cast("XTheme", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
