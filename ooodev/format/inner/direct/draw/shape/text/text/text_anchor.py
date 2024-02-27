from __future__ import annotations
from typing import cast, overload, TypeVar, Type, Any, Tuple
import uno
from ooo.dyn.drawing.text_horizontal_adjust import TextHorizontalAdjust
from ooo.dyn.drawing.text_vertical_adjust import TextVerticalAdjust

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.utils import props as mProps
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind


_TTextAnchor = TypeVar("_TTextAnchor", bound="TextAnchor")


class TextAnchor(StyleBase):
    """
    This class represents the text spacing of an object that supports ``com.sun.star.drawing.TextProperties``.

    .. versionadded:: 0.17.5
    """

    def __init__(
        self,
        anchor_point: ShapeBasePointKind | None = None,
        full_width: bool | None = None,
    ) -> None:
        """
        Constructor.

        Args:
            anchor_point (ShapeBasePointKind, optional): Anchor Point.
            full_width (bool, optional): Full Width. Defaults to None.

        Returns:
            None:

        Note:
            ``full_width`` applies when ``anchor_point`` is ``None`` or
            ``ShapeBasePointKind.TOP_CENTER`` or ``ShapeBasePointKind.CENTER``
            or ``ShapeBasePointKind.BOTTOM_CENTER``.
        """
        super().__init__()
        self._set_anchor(anchor_point, full_width)

    def _get_is_full_width(self) -> bool:
        horz = cast(TextHorizontalAdjust, self._get("TextHorizontalAdjust"))
        if horz is None:
            return False
        return horz == TextHorizontalAdjust.BLOCK

    def _get_anchor_point(self) -> ShapeBasePointKind | None:
        horz = cast(TextHorizontalAdjust, self._get("TextHorizontalAdjust"))
        vert = cast(TextVerticalAdjust, self._get("TextVerticalAdjust"))
        if horz is None or vert is None:
            return None
        if horz == TextHorizontalAdjust.LEFT:
            if vert == TextVerticalAdjust.TOP:
                return ShapeBasePointKind.TOP_LEFT
            elif vert == TextVerticalAdjust.CENTER:
                return ShapeBasePointKind.CENTER_LEFT
            elif vert == TextVerticalAdjust.BOTTOM:
                return ShapeBasePointKind.BOTTOM_LEFT
        elif horz == TextHorizontalAdjust.CENTER:
            if vert == TextVerticalAdjust.TOP:
                return ShapeBasePointKind.TOP_CENTER
            elif vert == TextVerticalAdjust.CENTER:
                return ShapeBasePointKind.CENTER
            elif vert == TextVerticalAdjust.BOTTOM:
                return ShapeBasePointKind.BOTTOM_CENTER
        elif horz == TextHorizontalAdjust.RIGHT:
            if vert == TextVerticalAdjust.TOP:
                return ShapeBasePointKind.TOP_RIGHT
            elif vert == TextVerticalAdjust.CENTER:
                return ShapeBasePointKind.CENTER_RIGHT
            elif vert == TextVerticalAdjust.BOTTOM:
                return ShapeBasePointKind.BOTTOM_RIGHT
        elif horz == TextHorizontalAdjust.BLOCK:
            if vert == TextVerticalAdjust.TOP:
                return ShapeBasePointKind.TOP_CENTER
            elif vert == TextVerticalAdjust.CENTER:
                return ShapeBasePointKind.CENTER
            elif vert == TextVerticalAdjust.BOTTOM:
                return ShapeBasePointKind.BOTTOM_CENTER
        raise mEx.NotSupportedError(f"Unknown anchor point: {horz}, {vert}")

    def _set_anchor(self, anchor_point: ShapeBasePointKind | None, full_width: bool | None) -> None:
        if anchor_point is None:
            if full_width is True:
                self._set("TextHorizontalAdjust", TextHorizontalAdjust.BLOCK)
                self._set("TextVerticalAdjust", TextVerticalAdjust.TOP)
                return
            self._remove("TextHorizontalAdjust")
            self._remove("TextVerticalAdjust")
            return

        if anchor_point == ShapeBasePointKind.TOP_LEFT:
            horz = TextHorizontalAdjust.LEFT
            vert = TextVerticalAdjust.TOP
        elif anchor_point == ShapeBasePointKind.TOP_CENTER:
            if full_width is True:
                horz = TextHorizontalAdjust.BLOCK
            else:
                horz = TextHorizontalAdjust.CENTER
            vert = TextVerticalAdjust.TOP
        elif anchor_point == ShapeBasePointKind.TOP_RIGHT:
            horz = TextHorizontalAdjust.RIGHT
            vert = TextVerticalAdjust.TOP
        elif anchor_point == ShapeBasePointKind.CENTER_LEFT:
            horz = TextHorizontalAdjust.LEFT
            vert = TextVerticalAdjust.CENTER
        elif anchor_point == ShapeBasePointKind.CENTER:
            if full_width is True:
                horz = TextHorizontalAdjust.BLOCK
            else:
                horz = TextHorizontalAdjust.CENTER
            vert = TextVerticalAdjust.CENTER
        elif anchor_point == ShapeBasePointKind.CENTER_RIGHT:
            horz = TextHorizontalAdjust.RIGHT
            vert = TextVerticalAdjust.CENTER
        elif anchor_point == ShapeBasePointKind.BOTTOM_LEFT:
            horz = TextHorizontalAdjust.LEFT
            vert = TextVerticalAdjust.BOTTOM
        elif anchor_point == ShapeBasePointKind.BOTTOM_CENTER:
            if full_width is True:
                horz = TextHorizontalAdjust.BLOCK
            else:
                horz = TextHorizontalAdjust.CENTER
            vert = TextVerticalAdjust.BOTTOM
        elif anchor_point == ShapeBasePointKind.BOTTOM_RIGHT:
            horz = TextHorizontalAdjust.RIGHT
            vert = TextVerticalAdjust.BOTTOM
        else:
            raise mEx.NotSupportedError(f"Unknown anchor point: {anchor_point}")
        self._set("TextHorizontalAdjust", horz)
        self._set("TextVerticalAdjust", vert)

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
    def from_obj(cls: Type[_TTextAnchor], obj: object) -> _TTextAnchor: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTextAnchor], obj: object, **kwargs) -> _TTextAnchor: ...

    @classmethod
    def from_obj(cls: Type[_TTextAnchor], obj: Any, **kwargs) -> _TTextAnchor:
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

        props = {"TextHorizontalAdjust", "TextVerticalAdjust"}

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
    def prop_anchor_point(self) -> ShapeBasePointKind | None:
        return self._get_anchor_point()

    @prop_anchor_point.setter
    def prop_anchor_point(self, value: ShapeBasePointKind | None) -> None:
        if value is None:
            self._remove("TextHorizontalAdjust")
            self._remove("TextVerticalAdjust")
            return
        full_width = self._get_is_full_width()
        self._set_anchor(value, full_width)

    @property
    def prop_full_width(self) -> bool | None:
        ap = self.prop_anchor_point
        if ap is None:
            return None
        return self._get_is_full_width()

    @prop_full_width.setter
    def prop_full_width(self, value: bool | None) -> None:
        if value is None:
            ap = self.prop_anchor_point
            if ap is None:
                return
        self._set_anchor(self.prop_anchor_point, value)

    # endregion Properties
