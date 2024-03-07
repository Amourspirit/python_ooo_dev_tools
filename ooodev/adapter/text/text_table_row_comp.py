from __future__ import annotations
import contextlib
from typing import TYPE_CHECKING, Tuple

from ooodev.adapter.beans.property_set_comp import PropertySetComp
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.text import TextTableRow  # service
    from com.sun.star.beans import PropertyValue  # struct
    from com.sun.star.text import TableColumnSeparator
    from com.sun.star.graphic import XGraphic

    # from com.sun.star.style.GraphicLocation import GraphicLocationProto  # type: ignore
    from ooo.dyn.style.graphic_location import GraphicLocation
    from ooodev.utils.color import Color
    from ooodev.units.unit_obj import UnitT


class TextTableRowComp(PropertySetComp):
    """
    Class for managing TextTableRow Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: TextTableRow) -> None:
        """
        Constructor

        Args:
            component (XTextTable): UNO Component that support ``com.sun.star.text.TextTableRow`` service.
        """
        PropertySetComp.__init__(self, component=component)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextTableRow",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> TextTableRow:
        """TextTableRow Component"""
        # pylint: disable=no-member
        return super().component  # type: ignore

    @property
    def row_interop_grab_bag(self) -> Tuple[PropertyValue, ...] | None:
        """
        Gets/Sets - Grab bag of row properties, used as a string-any map for interop purposes.

        This property is intentionally not handled by the ODF filter. Any member that should be handled there should be first moved out from this grab bag to a separate property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.RowInteropGrabBag
        return None

    @row_interop_grab_bag.setter
    def row_interop_grab_bag(self, value: Tuple[PropertyValue, ...]) -> None:
        with contextlib.suppress(AttributeError):
            self.component.RowInteropGrabBag = value

    @property
    def table_column_separators(self) -> Tuple[TableColumnSeparator, ...]:
        """
        Gets/Sets - contains the description of the columns in the table row.
        """
        return self.component.TableColumnSeparators

    @table_column_separators.setter
    def table_column_separators(self, value: Tuple[TableColumnSeparator, ...]) -> None:
        self.component.TableColumnSeparators = value

    @property
    def back_color(self) -> Color:
        """
        Gets/Sets the color of the background.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return self.component.BackColor  # type: ignore

    @back_color.setter
    def back_color(self, value: Color) -> None:
        self.component.BackColor = value  # type: ignore

    @property
    def back_graphic(self) -> XGraphic | None:
        """
        Gets/Sets the graphic of the background.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.BackGraphic
        return None

    @back_graphic.setter
    def back_graphic(self, value: XGraphic) -> None:
        with contextlib.suppress(AttributeError):
            self.component.BackGraphic = value

    @property
    def back_graphic_filter(self) -> str:
        """
        Gets/Sets - the name of the file filter of a background graphic.
        """
        return self.component.BackGraphicFilter

    @back_graphic_filter.setter
    def back_graphic_filter(self, value: str) -> None:
        self.component.BackGraphicFilter = value

    @property
    def back_graphic_location(self) -> GraphicLocation:
        """
        Gets/Sets the position of the background graphic.

        Hint:
            - ``GraphicLocation`` can be imported from ``ooo.dyn.style.graphic_location``
        """
        return self.component.BackGraphicLocation  # type: ignore

    @back_graphic_location.setter
    def back_graphic_location(self, value: GraphicLocation) -> None:
        self.component.BackGraphicLocation = value  # type: ignore

    @property
    def back_graphic_url(self) -> str:
        """
        Gets/Sets the URL of a background graphic.

        Note the new behavior since it this was deprecated: This property can only be set and only external URLs are supported (no more vnd.sun.star.GraphicObject scheme). When an URL is set, then it will load the graphic and set the BackGraphic property.
        """
        return self.component.BackGraphicURL

    @back_graphic_url.setter
    def back_graphic_url(self, value: str) -> None:
        self.component.BackGraphicURL = value

    @property
    def back_transparent(self) -> bool:
        """
        If ``True``, the background color value in ``back_color`` is not visible.
        """
        return self.component.BackTransparent

    @back_transparent.setter
    def back_transparent(self, value: bool) -> None:
        self.component.BackTransparent = value

    @property
    def has_text_changes_only(self) -> bool | None:
        """
        Gets/Sets - If ``True``, the table row wasn't deleted or inserted with its tracked cell content.

        **Since** LibreOffice 7.2

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.HasTextChangesOnly
        return None

    @has_text_changes_only.setter
    def has_text_changes_only(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.HasTextChangesOnly = value

    @property
    def height(self) -> UnitMM100:
        """
        Gets/Sets the height of the table row.


        When setting the height, the ``is_auto_height`` property must be ``False`` for the height to have an effect.
        Setting the height can be done with a ``UnitT`` or an integer in ``1/100th mm`` units.

        Returns:
            UnitMM100: Height

        Note:
            ``is_auto_height`` must be ``False`` for height to have an effect.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``
        """
        return UnitMM100(self.component.Height)

    @height.setter
    def height(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.component.Height = val.value

    @property
    def is_auto_height(self) -> bool:
        """
        Gets/Sets - If the value of this property is ``True``, the height of the table row depends on the content of the table cells.
        """
        return self.component.IsAutoHeight

    @is_auto_height.setter
    def is_auto_height(self, value: bool) -> None:
        self.component.IsAutoHeight = value

    @property
    def is_split_allowed(self) -> bool | None:
        """
        If ``True``, the row is allowed to be split at page or column breaks.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.IsSplitAllowed
        return None

    @is_split_allowed.setter
    def is_split_allowed(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.IsSplitAllowed = value

    # endregion Properties
