from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.form.forms_partial import FormsPartial

if TYPE_CHECKING:
    from com.sun.star.form import Forms


class FormsComp(ComponentBase, FormsPartial):
    """
    Class for managing Forms Service.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.form.Forms`` service.
        """

        ComponentBase.__init__(self, component)
        FormsPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.Forms",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> Forms:
        """Forms Component"""
        # pylint: disable=no-member
        return cast("Forms", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
