from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from com.sun.star.sheet import XArrayFormulaTokens

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.sheet import FormulaToken
    from ooodev.utils.type_var import UnoInterface


class ArrayFormulaTokensPartial:
    """
    Partial Class for XArrayFormulaTokens.
    """

    def __init__(
        self,
        component: XArrayFormulaTokens,
        interface: UnoInterface | None = XArrayFormulaTokens,
    ) -> None:
        """
        Constructor

        Args:
            component (XArrayFormulaTokens): UNO Component that implements ``com.sun.star.sheet.XArrayFormulaTokens``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XArrayFormulaTokens``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XArrayFormulaTokens
    def get_array_tokens(self) -> Tuple[FormulaToken, ...]:
        """
        Returns the array formula as sequence of tokens.
        """
        return self.__component.getArrayTokens()

    def set_array_tokens(self, *tokens: FormulaToken) -> None:
        """
        Sets the array formula as sequence of tokens.

        Args:
            tokens (FormulaToken): One or more formula as sequence of tokens.
        """
        self.__component.setArrayTokens(tokens)

    # endregion XArrayFormulaTokens
