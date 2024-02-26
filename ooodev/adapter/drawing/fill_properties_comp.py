from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase

from ooodev.adapter.drawing.fill_properties_partial import FillPropertiesPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service
    from com.sun.star.lang import XComponent


class FillPropertiesComp(ComponentBase, FillPropertiesPartial):
    """
    Class for managing FillProperties Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XComponent) -> None:
        """
        Constructor

        Args:
            component (XComponent): UNO component that implements ``com.sun.star.drawing.FillProperties`` service.
        """
        ComponentBase.__init__(self, component)
        FillPropertiesPartial.__init__(self, component=component)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by FillPropertiesPartial
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> FillProperties:
        """FillProperties Component"""
        # pylint: disable=no-member
        return cast("FillProperties", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
