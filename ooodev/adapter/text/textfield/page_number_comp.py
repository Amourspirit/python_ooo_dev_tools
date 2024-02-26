from __future__ import annotations

from typing import cast, TYPE_CHECKING
import uno

from ooo.dyn.style.numbering_type import NumberingTypeEnum
from ooo.dyn.text.page_number_type import PageNumberType
from ooodev.adapter.text.text_field_comp import TextFieldComp

if TYPE_CHECKING:
    from com.sun.star.text.textfield import PageNumber  # service
    from com.sun.star.text import XTextField


class PageNumberComp(TextFieldComp):
    """
    Class for managing PageNumber Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextField) -> None:
        """
        Constructor

        Args:
            component (XTextField): UNO PageNumber Component that supports ``com.sun.star.text.textfield.PageNumber`` service.
        """

        TextFieldComp.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.textfield.PageNumber",)

    # endregion Overrides

    # region Properties
    @property
    def offset(self) -> int:
        """Gets or sets an offset value to show a different page number."""
        return self.component.Offset

    @offset.setter
    def offset(self, value: int) -> None:
        self.component.Offset = value

    @property
    def number_type(self) -> NumberingTypeEnum:
        """Gets or sets the numbering type."""
        return NumberingTypeEnum(self.component.NumberingType)

    @number_type.setter
    def number_type(self, value: NumberingTypeEnum) -> None:
        self.component.NumberingType = value.value

    @property
    def sub_type(self) -> PageNumberType:
        """Gets or sets the numbering type."""
        return PageNumberType(self.component.SubType)

    @sub_type.setter
    def sub_type(self, value: PageNumberType) -> None:
        self.component.SubType = value.value  # type: ignore

    @property
    def user_text(self) -> str:
        """
        Gets or sets the user text.

        If the user text string is set then it is displayed when the value of
        NumberingType is set to ``com.sun.star.style.NumberingType.CHAR_SPECIAL``
        """
        return self.component.UserText

    @user_text.setter
    def user_text(self, value: str) -> None:
        self.component.UserText = value

    if TYPE_CHECKING:

        @property
        def component(self) -> PageNumber:
            """PageNumber Component"""
            # pylint: disable=no-member
            return cast("PageNumber", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
