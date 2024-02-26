"""
Module for Image Crop (``GraphicCrop``) struct

.. versionadded:: 0.9.0
"""

# region imports
from __future__ import annotations
from typing import Any, Tuple, Type, cast, overload, TypeVar

from ooo.dyn.util.cell_protection import CellProtection

from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.props.struct_cell_protection_props import StructCellProtectionProps
from ooodev.format.inner.direct.structs.struct_base import StructBase

# endregion imports

_TCellProtectionStruct = TypeVar(name="_TCellProtectionStruct", bound="CellProtectionStruct")


class CellProtectionStruct(StructBase):
    """
    Cell Protection struct.

    Any properties starting with ``prop_`` set or get current instance values.
    """

    # region init

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
        """

        super().__init__()
        self.prop_hide_all = hide_all
        self.prop_hide_formula = hide_formula
        self.prop_protected = protected
        self.prop_hide_print = hide_print

    # endregion init

    # region internal methods
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "CellProtection"
        return self._property_name

    # endregion internal methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        obj2 = None
        if isinstance(oth, CellProtectionStruct):
            obj2 = oth.get_uno_struct()
        if getattr(oth, "typeName", None) == "com.sun.star.util.CellProtection":
            obj2 = cast(CellProtection, oth)
        if obj2:
            obj1 = self.get_uno_struct()
            return (
                obj1.IsLocked == obj2.IsLocked
                and obj1.IsFormulaHidden == obj2.IsFormulaHidden
                and obj1.IsHidden == obj2.IsHidden
                and obj1.IsPrintHidden == obj2.IsPrintHidden
            )
        return NotImplemented

    # endregion dunder methods

    # region methods
    def get_uno_struct(self) -> CellProtection:
        """
        Gets UNO ``CellProtection`` from instance.

        Returns:
            Size: ``CellProtection`` instance
        """
        return CellProtection(
            IsLocked=self._get(self._props.protected),
            IsFormulaHidden=self._get(self._props.hide_formula),
            IsHidden=self._get(self._props.hide_all),
            IsPrintHidden=self._get(self._props.hide_print),
        )

    # endregion methods

    # region overrides methods

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle", "com.sun.star.table.CellProperties")
        return self._supported_services_values

    # region apply()

    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        if not mProps.Props.has(obj, self._get_property_name()):
            self._print_not_valid_srv("apply")
            return

        struct = self.get_uno_struct()
        props = {self._get_property_name(): struct}
        super().apply(obj=obj, override_dv=props)

    # endregion apply()

    # endregion overrides methods

    # region static methods

    # region from_uno_struct()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TCellProtectionStruct], value: CellProtection) -> _TCellProtectionStruct: ...

    @overload
    @classmethod
    def from_uno_struct(
        cls: Type[_TCellProtectionStruct], value: CellProtection, **kwargs
    ) -> _TCellProtectionStruct: ...

    @classmethod
    def from_uno_struct(cls: Type[_TCellProtectionStruct], value: CellProtection, **kwargs) -> _TCellProtectionStruct:
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

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TCellProtectionStruct], obj: Any) -> _TCellProtectionStruct: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TCellProtectionStruct], obj: Any, **kwargs) -> _TCellProtectionStruct: ...

    @classmethod
    def from_obj(cls: Type[_TCellProtectionStruct], obj: Any, **kwargs) -> _TCellProtectionStruct:
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
            struct = cast(CellProtection, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError as e:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property") from e

        return cls.from_uno_struct(struct, **kwargs)

    # endregion from_obj()

    # endregion static methods

    # region Format Properties
    @property
    def hide_all(self: _TCellProtectionStruct) -> _TCellProtectionStruct:
        """Gets instance with hide all value set."""
        cp = self.copy()
        cp.prop_hide_all = True
        return cp

    @property
    def hide_formula(self: _TCellProtectionStruct) -> _TCellProtectionStruct:
        """Gets instance with hide formula value set."""
        cp = self.copy()
        cp.prop_hide_formula = True
        return cp

    @property
    def protected(self: _TCellProtectionStruct) -> _TCellProtectionStruct:
        """Gets instance with protected value set."""
        cp = self.copy()
        cp.prop_protected = True
        return cp

    @property
    def hide_print(self: _TCellProtectionStruct) -> _TCellProtectionStruct:
        """Gets instance with hide print value set."""
        cp = self.copy()
        cp.prop_hide_print = True
        return cp

    # endregion Format Properties

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STRUCT
        return self._format_kind_prop

    @property
    def prop_hide_all(self) -> bool:
        """Gets/Sets Hide all value"""
        return self._get(self._props.hide_all)

    @prop_hide_all.setter
    def prop_hide_all(self, value: bool) -> None:
        self._set(self._props.hide_all, value)

    @property
    def prop_hide_formula(self) -> bool:
        """Gets/Sets Hide Formula value"""
        return self._get(self._props.hide_formula)

    @prop_hide_formula.setter
    def prop_hide_formula(self, value: bool) -> None:
        self._set(self._props.hide_formula, value)

    @property
    def prop_protected(self) -> bool:
        """Gets/Sets protected value"""
        return self._get(self._props.protected)

    @prop_protected.setter
    def prop_protected(self, value: bool) -> None:
        self._set(self._props.protected, value)

    @property
    def prop_hide_print(self) -> bool:
        """Gets/Sets Hide Print value"""
        return self._get(self._props.hide_print)

    @prop_hide_print.setter
    def prop_hide_print(self, value: bool) -> None:
        self._set(self._props.hide_print, value)

    @property
    def _props(self) -> StructCellProtectionProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = StructCellProtectionProps(
                hide_all="IsHidden", hide_formula="IsFormulaHidden", protected="IsLocked", hide_print="IsPrintHidden"
            )
        return self._props_internal_attributes

    # endregion Properties
