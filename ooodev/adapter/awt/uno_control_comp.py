from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.awt.control_partial import ControlPartial
from ooodev.adapter.awt.window_partial import WindowPartial
from ooodev.adapter.awt.view_partial import ViewPartial
from ooodev.units.size_pos_px import SizePosPX

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControl


class UnoControlComp(ComponentBase, WindowPartial, ViewPartial, ControlPartial):
    """
    Class for managing UnoControl Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: UnoControl) -> None:
        """
        Constructor

        Args:
            component (UnoControl): UNO Component that implements ``com.sun.star.awt.UnoControl`` service.
        """

        ComponentBase.__init__(self, component)
        WindowPartial.__init__(self, component=self.component, interface=None)
        ViewPartial.__init__(self, component=self.component, interface=None)
        ControlPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControl",)

    # region override XWindow
    @override
    def get_pos_size(self) -> SizePosPX:  # type: ignore
        """
        Gets the outer bounds of the window.
        """
        rect = self.component.getPosSize()
        return SizePosPX.from_unit_val(rect.X, rect.Y, rect.Width, rect.Height)

    # endregion override XWindow

    # endregion Overrides

    # region Properties

    @property
    @override
    def component(self) -> UnoControl:
        """UnoControl Component"""
        # pylint: disable=no-member
        return cast("UnoControl", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
