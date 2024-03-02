from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_content_comp import TextContentComp


if TYPE_CHECKING:
    from com.sun.star.text import TextField  # service
    from com.sun.star.text import XTextField


class TextFieldComp(TextContentComp, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing TextField Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextField) -> None:
        """
        Constructor

        Args:
            component (XTextField): UNO TextField Component that supports ``com.sun.star.text.TextField`` service.
        """
        # pylint: disable=no-member
        TextContentComp.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextField",)

    # endregion Overrides

    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> TextField:
            """TextField Component"""
            # pylint: disable=no-member
            return cast("TextField", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
