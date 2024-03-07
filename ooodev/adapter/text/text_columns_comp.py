from __future__ import annotations
from typing import cast, TYPE_CHECKING
import contextlib
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.text.text_columns_partial import TextColumnsPartial

if TYPE_CHECKING:
    from com.sun.star.text import TextColumns  # service
    from com.sun.star.text import XTextColumns
    from ooodev.utils.color import Color
    from ooo.dyn.style.vertical_alignment import VerticalAlignment


class TextColumnsComp(ComponentBase, TextColumnsPartial):
    """
    Class for managing TextColumns Component.

    Provides methods to access columns via index and to insert and remove columns.
    """

    # this class is very similar to ooodev.adapter.table.table_columns_comp.TableColumnsComp
    # don't get them confused.
    # pylint: disable=unused-argument

    def __init__(self, component: XTextColumns) -> None:
        """
        Constructor

        Args:
            component (XCell): UNO Component that implements ``com.sun.star.text.TextColumns`` service.
        """
        ComponentBase.__init__(self, component)
        TextColumnsPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextColumns",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> TextColumns:
        """Table columns Component"""
        # pylint: disable=no-member
        return cast("TextColumns", self._ComponentBase__get_component())  # type: ignore

    @property
    def automatic_distance(self) -> int:
        """
        Gets/Sets the distance between the columns.

        It is valid if the property IsAutomatic is set. Half of this distance is set to the left and right margins of all columns, except for the left margin of the first column, and the right margin of the last column.
        """
        return self.component.AutomaticDistance

    @automatic_distance.setter
    def automatic_distance(self, value: int) -> None:
        self.component.AutomaticDistance = value

    @property
    def is_automatic(self) -> bool:
        """
        Gets whether the columns all have equal width.

        This flag is set if XTextColumns.setColumnCount() is called and it is reset if XTextColumns.setColumns() is called.
        """
        return self.component.IsAutomatic

    @property
    def separator_line_color(self) -> Color:
        """
        Gets/Sets the color of the separator lines between the columns.

        Returns:
            ~ooodev.utils.color.Color: Color of the separator lines between the columns.
        """
        return self.component.SeparatorLineColor  # type: ignore

    @separator_line_color.setter
    def separator_line_color(self, value: Color) -> None:
        self.component.SeparatorLineColor = value  # type: ignore

    @property
    def separator_line_is_on(self) -> bool:
        """
        Gets/Sets whether separator lines are on.
        """
        return self.component.SeparatorLineIsOn

    @separator_line_is_on.setter
    def separator_line_is_on(self, value: bool) -> None:
        self.component.SeparatorLineIsOn = value

    @property
    def separator_line_relative_height(self) -> int:
        """
        Gets/Sets the relative height of the separator lines between the columns.
        """
        return self.component.SeparatorLineRelativeHeight

    @separator_line_relative_height.setter
    def separator_line_relative_height(self, value: int) -> None:
        self.component.SeparatorLineRelativeHeight = value

    @property
    def separator_line_style(self) -> int | None:
        """
        Gets/Sets the style of the separator lines between the columns.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.SeparatorLineStyle
        return None

    @separator_line_style.setter
    def separator_line_style(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.component.SeparatorLineStyle = value

    @property
    def separator_line_vertical_alignment(self) -> VerticalAlignment:
        """
        Gets/Sets the vertical alignment of the separator lines between the columns.

        Returns:
            VerticalAlignment: Vertical alignment of the separator lines between the columns.

        Hint:
            - ``VerticalAlignment`` can be imported from ``ooo.dyn.style.vertical_alignment``
        """
        return self.component.SeparatorLineVerticalAlignment  # type: ignore

    @separator_line_vertical_alignment.setter
    def separator_line_vertical_alignment(self, value: VerticalAlignment) -> None:
        self.component.SeparatorLineVerticalAlignment = value  # type: ignore

    @property
    def separator_line_width(self) -> int:
        """
        Gets/Sets the width of the separator lines between the columns.
        """
        return self.component.SeparatorLineWidth

    @separator_line_width.setter
    def separator_line_width(self, value: int) -> None:
        self.component.SeparatorLineWidth = value

    # endregion Properties
