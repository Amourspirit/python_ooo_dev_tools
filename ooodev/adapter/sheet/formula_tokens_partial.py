from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from com.sun.star.sheet import XFormulaTokens

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.sheet import FormulaToken
    from ooodev.utils.type_var import UnoInterface


class FormulaTokensPartial:
    """
    Partial Class for XFormulaTokens.
    """

    def __init__(
        self,
        component: XFormulaTokens,
        interface: UnoInterface | None = XFormulaTokens,
    ) -> None:
        """
        Constructor

        Args:
            component (XFormulaTokens): UNO Component that implements ``com.sun.star.sheet.XFormulaTokens``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFormulaTokens``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XFormulaTokens
    def get_tokens(self) -> Tuple[FormulaToken, ...]:
        """
        returns the formula as sequence of tokens.
        """
        return self.__component.getTokens()

    def setTokens(self, *tokens: FormulaToken) -> None:
        """
        Sets the formula as sequence of tokens.

        Args:
            tokens (FormulaToken): One or more formula as sequence of tokens.
        """
        self.__component.setTokens(tokens)

    # endregion XFormulaTokens
