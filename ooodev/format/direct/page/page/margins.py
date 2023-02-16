"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar, NamedTuple

import uno
from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from .....utils.unit_convert import UnitConvert

_TMargins = TypeVar(name="_TMargins", bound="Margins")


class MarginProps(NamedTuple):
    left: str
    right: str
    top: str
    bottom: str
    gutter: str


class Margins(StyleBase):
    """
    Fill Transparency

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        gutter: float | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Left Margin Value in ``mm`` units.
            right (float, optional): Right Margin Value in ``mm`` units.
            top (float, optional): Top Margin Value in ``mm`` units.
            bottom (float, optional): Bottom Margin Value in ``mm`` units.
            gutter (float, optional): Gutter Margin Value in ``mm`` units.
        """

        super().__init__()
        self.prop_left = left
        self.prop_right = right
        self.prop_top = top
        self.prop_bottom = bottom
        self.prop_gutter = gutter

    # region Internal Methods
    def _check(self, value: float | None, name: str) -> None:
        if value is None:
            return
        if value < 0.0:
            raise ValueError(f'"{name}" parameter must be a positive value')

    # endregion Internal Methods

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Margins):
            result = True
            for attr in self._props:
                result = result and self._get(attr) == oth._get(attr)
            return result
        return NotImplemented

    # endregion dunder methods

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    @classmethod
    def from_obj(cls: Type[_TMargins], obj: object) -> _TMargins:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Returns:
            Margins: Instance that represents object margins.
        """
        # this nu is only used to get Property Name

        inst = super(Margins, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        for attrib in inst._props:
            inst._set(attrib, mProps.Props.get(obj, attrib))
        return inst

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PAGE

    @property
    def prop_top(self) -> float | None:
        """Gets/Sets Top value"""
        pv = cast(int, self._get(self._props.top))
        if pv is None:
            return None
        return UnitConvert.convert_mm100_mm(pv)

    @prop_top.setter
    def prop_top(self, value: float | None) -> None:
        self._check(value, "top")
        if value is None:
            self._remove(self._props.top)
            return
        self._set(self._props.top, UnitConvert.convert_mm_mm100(value))

    @property
    def prop_bottom(self) -> float | None:
        """Gets/Sets Bottom value"""
        pv = cast(int, self._get(self._props.bottom))
        if pv is None:
            return None
        return UnitConvert.convert_mm100_mm(pv)

    @prop_bottom.setter
    def prop_bottom(self, value: float | None) -> None:
        self._check(value, "bottom")
        if value is None:
            self._remove(self._props.bottom)
            return
        self._set(self._props.bottom, UnitConvert.convert_mm_mm100(value))

    @property
    def prop_left(self) -> float | None:
        """Gets/Sets Left value"""
        pv = cast(int, self._get(self._props.left))
        if pv is None:
            return None
        return UnitConvert.convert_mm100_mm(pv)

    @prop_left.setter
    def prop_left(self, value: float | None) -> None:
        self._check(value, "left")
        if value is None:
            self._remove(self._props.left)
            return
        self._set(self._props.left, UnitConvert.convert_mm_mm100(value))

    @property
    def prop_right(self) -> float | None:
        """Gets/Sets Right value"""
        pv = cast(int, self._get(self._props.right))
        if pv is None:
            return None
        return UnitConvert.convert_mm100_mm(pv)

    @prop_right.setter
    def prop_right(self, value: float | None) -> None:
        self._check(value, "right")
        if value is None:
            self._remove(self._props.right)
            return
        self._set(self._props.right, UnitConvert.convert_mm_mm100(value))

    @property
    def prop_gutter(self) -> float | None:
        """Gets/Sets Gutter value"""
        pv = cast(int, self._get(self._props.gutter))
        if pv is None:
            return None
        return UnitConvert.convert_mm100_mm(pv)

    @prop_gutter.setter
    def prop_gutter(self, value: float | None) -> None:
        self._check(value, "gutter")
        if value is None:
            self._remove(self._props.gutter)
            return
        self._set(self._props.gutter, UnitConvert.convert_mm_mm100(value))

    @property
    def _props(self) -> MarginProps:
        try:
            return self._props_margins
        except AttributeError:
            self._props_margins = MarginProps(
                left="LeftMargin", right="RightMargin", top="TopMargin", bottom="BottomMargin", gutter="GutterMargin"
            )
        return self._props_margins

    @static_prop
    def default() -> Margins:  # type: ignore[misc]
        """Gets Margin Default. Static Property."""
        try:
            return Margins._DEFAULT_INST
        except AttributeError:
            inst = Margins()
            for attrib in inst._props:
                inst._set(attrib, 2000)
            inst._set(inst._props.gutter, 0)
            inst._is_default_inst = True
            Margins._DEFAULT_INST = inst
        return Margins._DEFAULT_INST
