from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial
from ooodev.utils.partial.model_prop_partial import ModelPropPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlModel


class UnoControlModelComp(ComponentBase, ModelPropPartial, UnoControlModelPartial):
    """
    Class for managing UnoControlDialog Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: UnoControlModel) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that implements ``com.sun.star.awt.UnoControlDialog`` service.
        """

        ComponentBase.__init__(self, component)
        ModelPropPartial.__init__(self, obj=component)  # must precede UnoControlModelPartial
        UnoControlModelPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlModel",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> UnoControlModel:
        """UnoControlModel Component"""
        # pylint: disable=no-member
        return cast("UnoControlModel", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
