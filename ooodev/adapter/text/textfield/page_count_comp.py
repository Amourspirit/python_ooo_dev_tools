from __future__ import annotations

from typing import cast, TYPE_CHECKING
import uno

# com.sun.star.style.NumberingType
from ooo.dyn.style.numbering_type import NumberingTypeEnum
from ooodev.adapter.text.text_field_comp import TextFieldComp

if TYPE_CHECKING:
    from com.sun.star.text.textfield import PageCount  # service
    from com.sun.star.text import XTextField


class PageCountComp(TextFieldComp):
    """
    Class for managing PageCount Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextField) -> None:
        """
        Constructor

        Args:
            component (XTextField): UNO PageCount Component that supports ``com.sun.star.text.textfield.PageCount`` service.
        """

        TextFieldComp.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.textfield.PageCount",)

    # endregion Overrides

    # region Properties
    @property
    def number_type(self) -> NumberingTypeEnum:
        """Gets or sets the numbering type."""
        return NumberingTypeEnum(self.component.NumberingType)

    @number_type.setter
    def number_type(self, value: NumberingTypeEnum) -> None:
        self.component.NumberingType = value.value

    if TYPE_CHECKING:

        @property
        def component(self) -> PageCount:
            """PageCount Component"""
            # pylint: disable=no-member
            return cast("PageCount", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
