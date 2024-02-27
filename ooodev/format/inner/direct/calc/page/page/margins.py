from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar, overload

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.page_margin_props import PageMarginProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.unit_obj import UnitT
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_convert import UnitConvert

_TMargins = TypeVar(name="_TMargins", bound="Margins")


class Margins(StyleBase):
    """
    Page Margins.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | UnitT | None = None,
        right: float | UnitT | None = None,
        top: float | UnitT | None = None,
        bottom: float | UnitT | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Left Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            right (float, optional): Right Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            top (float, optional): Top Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.
            bottom (float, optional): Bottom Margin Value in ``mm`` units  or :ref:`proto_unit_obj`.

        Returns:
            None:
        """

        super().__init__()
        self.prop_left = left
        self.prop_right = right
        self.prop_top = top
        self.prop_bottom = bottom

    # region Internal Methods
    def _check(self, value: float | UnitT | None, name: str) -> None:
        if value is None:
            return
        try:
            val = value.get_value_mm()  # type: ignore
        except AttributeError:
            val = float(value)  # type: ignore
        if val < 0.0:
            raise ValueError(f'"{name}" parameter must be a positive value')

    # endregion Internal Methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Margins):
            result = True
            for attr in self._props:
                if attr:
                    result = result and self._get(attr) == oth._get(attr)
            return result
        return NotImplemented

    # endregion dunder methods

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TMargins], obj: Any) -> _TMargins: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TMargins], obj: Any, **kwargs) -> _TMargins: ...

    @classmethod
    def from_obj(cls: Type[_TMargins], obj: Any, **kwargs) -> _TMargins:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Returns:
            Margins: Instance that represents object margins.
        """
        # this nu is only used to get Property Name

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        for attrib in inst._props:
            if attrib:
                inst._set(attrib, mProps.Props.get(obj, attrib))
        return inst

    # endregion from_obj()
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PAGE | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_top(self) -> UnitMM | None:
        """Gets/Sets Top value"""
        pv = cast(int, self._get(self._props.top))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_top.setter
    def prop_top(self, value: float | UnitT | None) -> None:
        self._check(value, "top")
        if value is None:
            self._remove(self._props.top)
            return
        try:
            self._set(self._props.top, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.top, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_bottom(self) -> UnitMM | None:
        """Gets/Sets Bottom value"""
        pv = cast(int, self._get(self._props.bottom))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_bottom.setter
    def prop_bottom(self, value: float | UnitT | None) -> None:
        self._check(value, "bottom")
        if value is None:
            self._remove(self._props.bottom)
            return
        try:
            self._set(self._props.bottom, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.bottom, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_left(self) -> UnitMM | None:
        """Gets/Sets Left value"""
        pv = cast(int, self._get(self._props.left))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_left.setter
    def prop_left(self, value: float | UnitT | None) -> None:
        self._check(value, "left")
        if value is None:
            self._remove(self._props.left)
            return
        try:
            self._set(self._props.left, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.left, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_right(self) -> UnitMM | None:
        """Gets/Sets Right value"""
        pv = cast(int, self._get(self._props.right))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_right.setter
    def prop_right(self, value: float | UnitT | None) -> None:
        self._check(value, "right")
        if value is None:
            self._remove(self._props.right)
            return
        try:
            self._set(self._props.right, value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set(self._props.right, UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def _props(self) -> PageMarginProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = PageMarginProps(
                left="LeftMargin", right="RightMargin", top="TopMargin", bottom="BottomMargin", gutter=""
            )
        return self._props_internal_attributes

    @property
    def default(self: _TMargins) -> _TMargins:  # type: ignore[misc]
        """Gets Margin Default. Static Property."""
        try:
            return self._DEFAULT_INST
        except AttributeError:
            inst = self.__class__(_cattribs=self._get_internal_cattribs())  # type: ignore
            for attrib in inst._props:
                if attrib:
                    inst._set(attrib, 2000)
            inst._is_default_inst = True
            self._DEFAULT_INST = inst
        return self._DEFAULT_INST
