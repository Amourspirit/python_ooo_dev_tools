from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from com.sun.star.sheet import XFormulaOpCodeMapper
from ooo.dyn.sheet.formula_language import FormulaLanguageEnum

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.sheet import FormulaToken
    from com.sun.star.sheet import FormulaOpCodeMapEntry
    from ooodev.utils.type_var import UnoInterface


class FormulaOpCodeMapperPartial:
    """
    Partial Class for XFormulaOpCodeMapper.
    """

    def __init__(
        self,
        component: XFormulaOpCodeMapper,
        interface: UnoInterface | None = XFormulaOpCodeMapper,
    ) -> None:
        """
        Constructor

        Args:
            component (XFormulaOpCodeMapper): UNO Component that implements ``com.sun.star.sheet.XFormulaOpCodeMapper``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFormulaOpCodeMapper``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XFormulaOpCodeMapper
    def get_available_mappings(self, language: FormulaLanguageEnum, groups: int) -> Tuple[FormulaOpCodeMapEntry, ...]:
        """
        Returns a sequence of map entries for all available elements of a given formula language.

        Each element of the formula language in parameter Language is mapped to a FormulaToken containing the internal OpCode used by the spreadsheet application in FormulaToken.OpCode and by contract maybe additional information in FormulaToken.Data. See getMappings() for more details.

        Args:
            language (FormulaLanguageEnum): The formula language.
            groups (int): Group of mappings to be returned, a bit mask of ``FormulaMapGroup`` constants.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``

        Returns:
            Tuple[FormulaOpCodeMapEntry, ...]: A sequence of map entries for all available elements of a given formula language.

        Note:
            - ``FormulaLanguageEnum`` can be imported from ``ooo.dyn.sheet.formula_language``.
            - ``FormulaMapGroup`` can be imported from ``com.sun.star.sheet``.
        """
        return self.__component.getAvailableMappings(int(language), groups)

    def get_mappings(self, language: FormulaLanguageEnum, *names: str) -> Tuple[FormulaToken, ...]:
        """
        Returns a sequence of tokens matching the input sequence of strings in order.

        Each string element in parameter Names according to the formula language in parameter Language is mapped to a FormulaToken
        containing the internal OpCode used by the spreadsheet application in FormulaToken.
        OpCode and by contract maybe additional information in FormulaToken Data.

        The order of the FormulaToken sequence returned matches the input order of the string sequence.

        An unknown Name string gets the OpCode value of OpCodeUnknown assigned.

        Additional information in FormulaToken. Data is returned for:

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``

        Returns:
            Tuple[FormulaToken, ...]: A sequence of tokens matching the input sequence of strings in order.

        Note:
            - ``FormulaLanguageEnum`` can be imported from ``ooo.dyn.sheet.formula_language``.
        """
        return self.__component.getMappings(names, int(language))

    @property
    def op_code_external(self) -> int:
        """
        Gets - OpCode value used for external Add-In functions.

        Needed to be able to identify which of the function names map to an Add-In implementation where this OpCode
        is used in the returned mapping and the programmatic name is available as additional information.
        """
        return self.__component.OpCodeExternal

    @property
    def op_code_unknown(self) -> int:
        """
        Gets - OpCode value used for unknown functions.

        Used to identify which of the function names queried with getMappings() are unknown to the implementation.
        """
        return self.__component.OpCodeUnknown

    # endregion XFormulaOpCodeMapper
