from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.units.size_px import SizePX
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.text_component_partial import TextComponentPartial
from ooodev.adapter.awt.layout_constrains_partial import LayoutConstrainsPartial
from ooodev.adapter.awt.text_layout_constrains_partial import TextLayoutConstrainsPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlEdit


class UnoControlEditComp(UnoControlComp, TextComponentPartial, LayoutConstrainsPartial, TextLayoutConstrainsPartial):

    def __init__(self, component: UnoControlEdit):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlEdit`` service.
        """
        UnoControlComp.__init__(self, component=component)
        TextComponentPartial.__init__(self, component=self.component, interface=None)
        LayoutConstrainsPartial.__init__(self, component=self.component, interface=None)
        TextLayoutConstrainsPartial.__init__(self, component=self.component, interface=None)

    @override
    def get_minimum_size(self) -> SizePX:  # type: ignore
        """
        Gets the minimum size for this component.

        Returns:
            SizePX: Minimum size in pixel units.


        Note:
            Use ``get_minimum_size_text_layout()`` to get the minimum size for a given number of columns and lines.
        """
        result = self.component.getMinimumSize()
        return SizePX.from_unit_val(result.Width, result.Height)

    @property
    @override
    def component(self) -> UnoControlEdit:
        """UnoControlEdit Component"""
        # pylint: disable=no-member
        return cast("UnoControlEdit", self._ComponentBase__get_component())  # type: ignore
