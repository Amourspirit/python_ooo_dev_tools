from __future__ import annotations
from typing import TYPE_CHECKING, overload, TypeVar, Type, Any, Tuple
import contextlib
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TSpacing = TypeVar("_TSpacing", bound="Spacing")


class Spacing(StyleBase):
    """
    This class represents the text spacing of an object that supports ``com.sun.star.drawing.TextProperties``.
    """

    def __init__(
        self,
        left: float | UnitT | None = None,
        right: float | UnitT | None = None,
        top: float | UnitT | None = None,
        bottom: float | UnitT | None = None,
    ) -> None:
        """
        Constructor.

        Args:
            left (float, UnitT, optional): Left spacing in MM units or ``UnitT``. Defaults to None.
            right (float, UnitT, optional): Right spacing in MM units or ``UnitT``. Defaults to None.
            top (float, UnitT, optional): Top spacing in MM units or ``UnitT``. Defaults to None.
            bottom (float, UnitT, optional): Bottom spacing in MM units or ``UnitT``. Defaults to None.
        """
        super().__init__()
        self.prop_left = left
        self.prop_right = right
        self.prop_top = top
        self.prop_bottom = bottom

    def _get_unit_mm_100(self, value: float | UnitT | None) -> int | None:
        if value is None:
            return None
        with contextlib.suppress(AttributeError):
            return value.get_value_mm100()  # type: ignore
        return UnitMM(value).get_value_mm100()  # type: ignore

    # region Overridden Methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.TextProperties",)
        return self._supported_services_values

    # endregion Overridden Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TSpacing], obj: object) -> _TSpacing: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSpacing], obj: object, **kwargs) -> _TSpacing: ...

    @classmethod
    def from_obj(cls: Type[_TSpacing], obj: Any, **kwargs) -> _TSpacing:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Spacing: New instance.
        """
        inst = cls(**kwargs)

        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to Line Properties")

        props = {"TextLeftDistance", "TextRightDistance", "TextUpperDistance", "TextLowerDistance"}

        def set_property(prop: str):
            value = mProps.Props.get(obj, prop, None)
            if value is not None:
                inst._set(prop, value)

        for prop in props:
            set_property(prop)
        return inst

    # endregion from_obj()

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.SHAPE
        return self._format_kind_prop

    @property
    def prop_left(self) -> UnitMM | None:
        pv = self._get("TextLeftDistance")
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_left.setter
    def prop_left(self, value: float | UnitT | None) -> None:
        val = self._get_unit_mm_100(value)
        if val is None:
            self._remove("TextLeftDistance")
            return
        self._set("TextLeftDistance", val)

    @property
    def prop_right(self) -> UnitMM | None:
        pv = self._get("TextRightDistance")
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_right.setter
    def prop_right(self, value: float | UnitT | None) -> None:
        val = self._get_unit_mm_100(value)
        if val is None:
            self._remove("TextRightDistance")
            return
        self._set("TextRightDistance", val)

    @property
    def prop_top(self) -> UnitMM | None:
        pv = self._get("TextUpperDistance")
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_top.setter
    def prop_top(self, value: float | UnitT | None) -> None:
        val = self._get_unit_mm_100(value)
        if val is None:
            self._remove("TextUpperDistance")
            return
        self._set("TextUpperDistance", val)

    @property
    def prop_bottom(self) -> UnitMM | None:
        pv = self._get("TextLowerDistance")
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_bottom.setter
    def prop_bottom(self, value: float | UnitT | None) -> None:
        val = self._get_unit_mm_100(value)
        if val is None:
            self._remove("TextLowerDistance")
            return
        self._set("TextLowerDistance", val)

    # endregion Properties
