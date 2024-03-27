from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.form.runtime.form_controller_partial import FormControllerPartial

if TYPE_CHECKING:
    from com.sun.star.form.runtime import XFormController
    from com.sun.star.form.runtime import FormController


class FormControllerComp(ComponentBase, FormControllerPartial):
    """
    Class for managing FormController.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XFormController) -> None:
        """
        Constructor

        Args:
            component (FormController): UNO Component that implements ``com.sun.star.form.FormController`` service.
        """

        ComponentBase.__init__(self, component)
        FormControllerPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.runtime.FormController",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> FormController:
        """FormController Component"""
        # pylint: disable=no-member
        return cast("FormController", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
