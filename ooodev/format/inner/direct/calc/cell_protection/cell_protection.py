from __future__ import annotations
from typing import Any, cast, overload, TypeVar, Type
import uno
from ooo.dyn.util.cell_protection import CellProtection as UnoCellProtection
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.direct.structs.cell_protection_struct import CellProtectionStruct

_TCellProtection = TypeVar("_TCellProtection", bound="CellProtection")


class CellProtection(CellProtectionStruct):
    """
    Cell Protection.

    Warning:
        Cell protection is only effective after the sheet is has been applied to is also been protected.

    .. seealso::

        - :ref:`help_calc_format_direct_cell_cell_protection`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self, hide_all: bool = False, protected: bool = False, hide_formula: bool = False, hide_print: bool = False
    ) -> None:
        """
        Constructor

        Args:
            hide_all (bool, optional): Specifies if all is hidden. Defaults to ``False``.
            protected (bool, optional): Specifies protected value. Defaults to ``False``.
            hide_formula (bool, optional): Specifies if the formula is hidden. Defaults to ``False``.
            hide_print (bool, optional): Specifies if the cell are to be omitted during print. Defaults to ``False``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_direct_cell_cell_protection`
        """
        super().__init__(hide_all=hide_all, protected=protected, hide_formula=hide_formula, hide_print=hide_print)

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TCellProtection], obj: Any) -> _TCellProtection: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TCellProtection], obj: Any, **kwargs) -> _TCellProtection: ...

    @classmethod
    def from_obj(cls: Type[_TCellProtection], obj: Any, **kwargs) -> _TCellProtection:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            CellProtectionStruct: ``CellProtectionStruct`` instance that represents ``obj`` CellProtection properties.
        """
        # this nu is only used to get Property Name
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()

        try:
            struct = cast(UnoCellProtection, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError as e:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property") from e

        result = cls.from_uno_struct(struct, **kwargs)
        result.set_update_obj(obj)
        return result

    # region from_uno_struct()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TCellProtection], value: UnoCellProtection) -> _TCellProtection: ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TCellProtection], value: UnoCellProtection, **kwargs) -> _TCellProtection: ...

    @classmethod
    def from_uno_struct(cls: Type[_TCellProtection], value: UnoCellProtection, **kwargs) -> _TCellProtection:
        """
        Converts a ``CellProtection`` instance to a ``CellProtectionStruct``.

        Args:
            value (CellProtection): UNO ``CellProtection``.

        Returns:
            CellProtectionStruct: ``CellProtectionStruct`` set with ``CellProtection`` properties.
        """
        inst = cls(**kwargs)
        inst._set(inst._props.protected, value.IsLocked)
        inst._set(inst._props.hide_formula, value.IsFormulaHidden)
        inst._set(inst._props.hide_all, value.IsHidden)
        inst._set(inst._props.hide_print, value.IsPrintHidden)
        return inst

    # endregion from_uno_struct()
    # endregion from_obj()
