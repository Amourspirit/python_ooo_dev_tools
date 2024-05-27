from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.sheet import XNamedRange
from ooo.dyn.sheet.named_range_flag import NamedRangeFlagEnum

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.container.named_partial import NamedPartial

if TYPE_CHECKING:
    from com.sun.star.table import CellAddress
    from ooodev.utils.type_var import UnoInterface


class NamedRangePartial(NamedPartial):
    """
    Partial Class for XNamedRange.
    """

    def __init__(self, component: XNamedRange, interface: UnoInterface | None = XNamedRange) -> None:
        """
        Constructor

        Args:
            component (XNamedRange): UNO Component that implements ``com.sun.star.sheet.XNamedRange``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNamedRange``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        NamedPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XNamedRange

    def get_content(self) -> str:
        """
        Returns the content of the named range.

        The content can be a reference to a cell or cell range or any formula expression.
        """
        return self.__component.getContent()

    def get_reference_position(self) -> CellAddress:
        """
        Returns the position in the document which is used as a base for relative references in the content.

        Returns:
            CellAddress: The position in the document.

        Note:
            ``CellAddress`` can be imported from ``com.sun.star.table``.
        """
        return self.__component.getReferencePosition()

    def get_type(self) -> int:
        """
        returns the type of the named range.

        This is a combination of flags as defined in NamedRangeFlag.

        Returns:
            NamedRangeFlagEnum: The type of the named range.

        Note:
            ``NamedRangeFlagEnum`` is a flags enum and can be imported from ``ooo.dyn.sheet.named_range_flag``.
        """
        # may return zero
        return self.__component.getType()

    def set_content(self, content: str) -> None:
        """
        sets the content of the named range.

        The content can be a reference to a cell or cell range or any formula expression.
        """
        self.__component.setContent(content)

    def set_reference_position(self, reference_position: CellAddress) -> None:
        """
        Sets the position in the document which is used as a base for relative references in the content.

        Args:
            reference_position (CellAddress): The position in the document.

        Returns:
            None:

        Note:
            ``CellAddress`` can be imported from ``com.sun.star.table``.
        """
        self.__component.setReferencePosition(reference_position)

    def set_type(self, type_enum: NamedRangeFlagEnum | int) -> None:
        """
        Sets the type of the named range.

        Args:
            type_enum (int | NamedRangeFlagEnum): The type of the named range.

        Returns:
            None:

        Note:
            This is a combination of flags as defined in NamedRangeFlag.

            ``NamedRangeFlagEnum`` is a flags enum and can be imported from ``ooo.dyn.sheet.named_range_flag``.
        """
        # can be zero
        self.__component.setType(int(type_enum))

    # endregion XNamedRange
