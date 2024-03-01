"""
This class is a partial class for ``com.sun.star.table.CellProperties`` service.
"""

from __future__ import annotations
from typing import TYPE_CHECKING
import uno


from ooodev.adapter.table.border_line_struct_comp import BorderLineStructComp
from ooodev.adapter.table.border_line2_struct_comp import BorderLine2StructComp
from ooodev.adapter.table.table_border2_struct_comp import TableBorder2StructComp
from ooodev.adapter.table.cell_properties_partial_props import CellPropertiesPartialProps


if TYPE_CHECKING:
    from com.sun.star.table import TableBorder2
    from com.sun.star.table import BorderLine
    from com.sun.star.table import BorderLine2


class CellProperties2PartialProps(CellPropertiesPartialProps):
    """
    Partial Class for CellProperties Service.

    This class is the same as ``CellPropertiesPartialProps`` except border are not optional.
    """

    # setting property override fix: https://stackoverflow.com/questions/1021464/how-to-call-a-property-of-the-base-class-if-this-property-is-being-overwritten-i
    # pylint: disable=unused-argument

    # region Properties

    @property
    def bottom_border2(self) -> BorderLine2StructComp:
        """
        Gets/Sets a description of the bottom border line of each cell.

        Preferred over ``bottom_border``.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2StructComp`` object.


        Returns:
            BorderLine2StructComp: Returns BorderLine2.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        comp = super().bottom_border2
        if comp is None:
            raise AttributeError("This property is not supported.")
        return comp

    @bottom_border2.setter
    def bottom_border2(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        super(CellProperties2PartialProps, self.__class__).bottom_border2.fset(self, value)  # type: ignore

    @property
    def diagonal_bltr(self) -> BorderLineStructComp:
        """
        Gets/Sets a description of the bottom left to top right diagonal line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLineComp`` object.

        Returns:
            BorderLineComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        comp = super().diagonal_bltr
        if comp is None:
            raise AttributeError("This property is not supported.")
        return comp

    @diagonal_bltr.setter
    def diagonal_bltr(self, value: BorderLine | BorderLineStructComp) -> None:
        super(CellProperties2PartialProps, self.__class__).diagonal_bltr.fset(self, value)  # type: ignore

    @property
    def diagonal_bltr2(self) -> BorderLine2StructComp:
        """
        Gets/Sets a description of the bottom left to top right diagonal line of each cell.

        Preferred over ``diagonal_bltr``.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2StructComp`` object.

        Returns:
            BorderLine2StructComp: Returns BorderLine2.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        comp = super().diagonal_bltr2
        if comp is None:
            raise AttributeError("This property is not supported.")
        return comp

    @diagonal_bltr2.setter
    def diagonal_bltr2(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        super(CellProperties2PartialProps, self.__class__).diagonal_bltr2.fset(self, value)  # type: ignore

    @property
    def diagonal_tlbr(self) -> BorderLineStructComp:
        """
        Gets/Sets a description of the top left to bottom right diagonal line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLineComp`` object.

        Returns:
            BorderLineComp: Returns BorderLine.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        comp = super().diagonal_tlbr
        if comp is None:
            raise AttributeError("This property is not supported.")
        return comp

    @diagonal_tlbr.setter
    def diagonal_tlbr(self, value: BorderLine | BorderLineStructComp) -> None:
        super(CellProperties2PartialProps, self.__class__).diagonal_tlbr.fset(self, value)  # type: ignore

    @property
    def diagonal_tlbr2(self) -> BorderLine2StructComp:
        """
        contains a description of the top left to bottom right diagonal line of each cell.

        Preferred over ``diagonal_tlbr``.

        Returns:
            BorderLine2StructComp: Returns BorderLine2.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        comp = super().diagonal_tlbr2
        if comp is None:
            raise AttributeError("This property is not supported.")
        return comp

    @diagonal_tlbr2.setter
    def diagonal_tlbr2(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        super(CellProperties2PartialProps, self.__class__).diagonal_tlbr2.fset(self, value)  # type: ignore

    @property
    def left_border2(self) -> BorderLine2StructComp:
        """
        Gets/Sets a description of the left border line of each cell.

        Preferred over ``left_border``.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2StructComp`` object.

        Returns:
            BorderLine2StructComp: Returns BorderLine2.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        comp = super().left_border2
        if comp is None:
            raise AttributeError("This property is not supported.")
        return comp

    @left_border2.setter
    def left_border2(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        super(CellProperties2PartialProps, self.__class__).left_border2.fset(self, value)  # type: ignore

    @property
    def right_border2(self) -> BorderLine2StructComp:
        """
        Gets/Sets a description of the right border line of each cell.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2StructComp`` object.

        Preferred over ``right_border``.

        Returns:
            BorderLine2StructComp: Returns BorderLine2.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        comp = super().right_border2
        if comp is None:
            raise AttributeError("This property is not supported.")
        return comp

    @right_border2.setter
    def right_border2(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        super(CellProperties2PartialProps, self.__class__).right_border2.fset(self, value)  # type: ignore

    @property
    def table_border2(self) -> TableBorder2StructComp:
        """
        Gets/Seta a description of the cell or cell range border.

        Preferred over ``table_border``.

        If used with a cell range, the top, left, right, and bottom lines are at the edges of the entire range, not at the edges of the individual cell.

        Setting value can be done with a ``TableBorder2`` or ``TableBorder2StructComp`` object.

        Returns:
            TableBorder2StructComp | None: Returns TableBorder2 or None if not supported.

        Hint:
            - ``TableBorder2`` can be imported from ``ooo.dyn.table.table_border2``
        """
        comp = super().table_border2
        if comp is None:
            raise AttributeError("This property is not supported.")
        return comp

    @table_border2.setter
    def table_border2(self, value: TableBorder2 | TableBorder2StructComp) -> None:
        super(CellProperties2PartialProps, self.__class__).table_border2.fset(self, value)  # type: ignore

    @property
    def top_border2(self) -> BorderLine2StructComp:
        """
        Gets/Sets a description of the top border line of each cell.

        Preferred over ``top_border``.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2StructComp`` object.

        Returns:
            BorderLine2StructComp: Returns BorderLine2.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        comp = super().top_border2
        if comp is None:
            raise AttributeError("This property is not supported.")
        return comp

    @top_border2.setter
    def top_border2(self, value: BorderLine2 | BorderLine2StructComp) -> None:
        super(CellProperties2PartialProps, self.__class__).top_border2.fset(self, value)  # type: ignore

    # endregion Properties
