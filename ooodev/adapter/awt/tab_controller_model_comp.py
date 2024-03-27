from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.awt.tab_controller_model_partial import TabControllerModelPartial

if TYPE_CHECKING:
    from com.sun.star.awt import XTabControllerModel


class TabControllerModelComp(ComponentBase, TabControllerModelPartial):
    """
    Class for managing XTabControllerModel Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTabControllerModel) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that implements ``com.sun.star.awt.XTabControllerModel`` interface.
        """

        ComponentBase.__init__(self, component)
        TabControllerModelPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XTabControllerModel:
        """XTabControllerModel Component"""
        # pylint: disable=no-member
        return cast("XTabControllerModel", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
