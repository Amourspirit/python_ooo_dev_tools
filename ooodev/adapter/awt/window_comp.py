from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno

from ooodev.adapter.awt.window_partial import WindowPartial
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.awt import XWindow


class WindowComp(ComponentBase, WindowPartial):
    """
    Class for managing Window Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XWindow) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.awt.XWindow`` interface.

        Returns:
            None:
        """

        ComponentBase.__init__(self, component)  # type: ignore
        WindowPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> XWindow:
        """XWindow Component"""
        # pylint: disable=no-member
        return cast("XWindow", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
