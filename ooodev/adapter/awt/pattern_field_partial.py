from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.awt import XPatternField

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PatternFieldPartial:
    """
    Partial class for XPatternField.
    """

    def __init__(self, component: XPatternField, interface: UnoInterface | None = XPatternField) -> None:
        """
        Constructor

        Args:
            component (XPatternField): UNO Component that implements ``com.sun.star.awt.XPatternField`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPatternField``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPatternField
    def get_masks(self) -> Tuple[str, str]:
        """
        Gets the currently set pattern mask.

        Returns:
            Tuple[str, str]: The currently set ``EditMask`` mask and ``LiteralMask`` mask.
        """
        # method has out args: EditMask, LiteralMask
        return self.__component.getMasks("", "")  # type: ignore

    def get_string(self) -> str:
        """
        Gets the currently set string value of the pattern field.
        """
        return self.__component.getString()

    def is_strict_format(self) -> bool:
        """
        Gets whether the format is currently checked during user input.
        """
        return self.__component.isStrictFormat()

    def set_masks(self, edit_mask: str, literal_mask: str) -> None:
        """
        Sets the pattern mask.
        """
        self.__component.setMasks(edit_mask, literal_mask)

    def set_strict_format(self, strict: bool) -> None:
        """
        Determines if the format is checked during user input.
        """
        self.__component.setStrictFormat(strict)

    def set_string(self, value: str) -> None:
        """
        Sets the string value of the pattern field.
        """
        self.__component.setString(value)

    # endregion XPatternField
