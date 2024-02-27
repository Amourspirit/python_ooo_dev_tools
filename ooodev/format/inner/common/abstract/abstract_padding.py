"""
Module for managing paragraph padding.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.units.unit_obj import UnitT
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_convert import UnitConvert
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.border_props import BorderProps

# endregion Imports
_TAbstractPadding = TypeVar(name="_TAbstractPadding", bound="AbstractPadding")


class AbstractPadding(StyleBase):
    """
    Abstract Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        left: float | UnitT | None = None,
        right: float | UnitT | None = None,
        top: float | UnitT | None = None,
        bottom: float | UnitT | None = None,
        all: float | UnitT | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitT, optional): Left (in ``mm`` units) or :ref:`proto_unit_obj`.
            right (float, UnitT, optional): Right (in ``mm`` units)  or :ref:`proto_unit_obj`.
            top (float, UnitT, optional): Top (in ``mm`` units)  or :ref:`proto_unit_obj`.
            bottom (float, UnitT,  optional): Bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
            all (float, UnitT, optional): Left, right, top, bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
                If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.

        Raises:
            ValueError: If any argument value is less than zero.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        def validate(val: float | UnitT | None) -> None:
            if val is not None:
                try:
                    value = val.get_value_mm()  # type: ignore
                except AttributeError:
                    value = float(val)  # type: ignore
                if value < 0:
                    raise ValueError("Values must be positive values")

        def set_val(key, value: float | UnitT) -> None:
            nonlocal init_vals
            if value is not None:
                try:
                    init_vals[key] = value.get_value_mm100()  # type: ignore
                except AttributeError:
                    init_vals[key] = UnitConvert.convert_mm_mm100(value)  # type: ignore

        validate(left)
        validate(right)
        validate(top)
        validate(bottom)
        validate(all)
        if all is None:
            for key, value in zip(self._props, (left, top, right, bottom)):
                if value is not None:
                    set_val(key, value)
        else:
            for key in self._props:
                set_val(key, all)

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CellStyle",
                "com.sun.star.style.ParagraphProperties",
            )
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: object) -> None:  # type: ignore
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractPadding], obj: object) -> _TAbstractPadding: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractPadding], obj: object, **kwargs) -> _TAbstractPadding: ...

    @classmethod
    def from_obj(cls: Type[_TAbstractPadding], obj: object, **kwargs) -> _TAbstractPadding:
        """
        Gets Border Padding instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            BorderPadding: BorderPadding that represents ``obj`` padding.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set(inst._props.left, int(mProps.Props.get(obj, inst._props.left)))
        inst._set(inst._props.right, int(mProps.Props.get(obj, inst._props.right)))
        inst._set(inst._props.top, int(mProps.Props.get(obj, inst._props.top)))
        inst._set(inst._props.bottom, int(mProps.Props.get(obj, inst._props.bottom)))
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()
    # endregion methods

    # region style methods
    def fmt_padding_all(self: _TAbstractPadding, value: float | UnitT | None) -> _TAbstractPadding:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (float, UnitT, optional): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_top = value
        cp.prop_bottom = value
        cp.prop_left = value
        cp.prop_right = value
        return cp

    def fmt_top(self: _TAbstractPadding, value: float | UnitT | None) -> _TAbstractPadding:
        """
        Gets a copy of instance with top side set or removed

        Args:
            value (float, Unit100MM, optional): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_top = value
        return cp

    def fmt_bottom(self: _TAbstractPadding, value: float | UnitT | None) -> _TAbstractPadding:
        """
        Gets a copy of instance with bottom side set or removed

        Args:
            value (float, UnitT, optional): Padding (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_bottom = value
        return cp

    def fmt_left(self: _TAbstractPadding, value: float | UnitT | None) -> _TAbstractPadding:
        """
        Gets a copy of instance with left side set or removed

        Args:
            value (float, UnitT, optional): Padding (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_left = value
        return cp

    def fmt_right(self: _TAbstractPadding, value: float | UnitT | None) -> _TAbstractPadding:
        """
        Gets a copy of instance with right side set or removed

        Args:
            value (float, UnitT, optional): Padding (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_right = value
        return cp

    # endregion style methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

    @property
    def prop_left(self) -> UnitMM | None:
        """Gets/Sets paragraph left padding (in mm units)."""
        pv = cast(int, self._get(self._props.left))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_left.setter
    def prop_left(self, value: float | UnitT | None):
        if value is None:
            self._remove(self._props.left)
            return
        try:
            self._set(self._props.left, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.left, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_right(self) -> UnitMM | None:
        """Gets/Sets paragraph right padding (in mm units)."""
        pv = cast(int, self._get(self._props.right))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_right.setter
    def prop_right(self, value: float | UnitT | None):
        if value is None:
            self._remove(self._props.right)
            return
        try:
            self._set(self._props.right, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.right, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_top(self) -> UnitMM | None:
        """Gets/Sets paragraph top padding (in mm units)."""
        pv = cast(int, self._get(self._props.top))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_top.setter
    def prop_top(self, value: float | UnitT | None):
        if value is None:
            self._remove(self._props.top)
            return
        try:
            self._set(self._props.top, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.top, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_bottom(self) -> UnitMM | None:
        """Gets/Sets paragraph bottom padding (in mm units)."""
        pv = cast(int, self._get(self._props.bottom))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_bottom.setter
    def prop_bottom(self, value: float | UnitT | None):
        if value is None:
            self._remove(self._props.bottom)
            return
        try:
            self._set(self._props.bottom, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.bottom, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def default(self: _TAbstractPadding) -> _TAbstractPadding:  # type: ignore[misc]
        """Gets Padding default."""
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(all=0.35, _cattribs=self._get_internal_cattribs())  # type: ignore
            self._default_inst._is_default_inst = True
        return self._default_inst

    @property
    def _props(self) -> BorderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = BorderProps(
                left="ParaLeftMargin", top="ParaTopMargin", right="ParaRightMargin", bottom="ParaBottomMargin"
            )
        return self._props_internal_attributes

    # endregion properties
