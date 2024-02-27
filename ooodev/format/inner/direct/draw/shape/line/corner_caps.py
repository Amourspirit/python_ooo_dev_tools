from __future__ import annotations
from typing import Tuple, overload, Any, TypeVar, Type
import uno

from ooo.dyn.drawing.line_joint import LineJoint
from ooo.dyn.drawing.line_cap import LineCap

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.utils import props as mProps


_TCornerCaps = TypeVar(name="_TCornerCaps", bound="CornerCaps")


class CornerCaps(StyleBase):
    """
    This class represents the line corner and cap styles of a shape.
    """

    def __init__(self, corner_style: LineJoint = LineJoint.ROUND, cap_style: LineCap = LineCap.BUTT) -> None:
        """
        Constructor.

        Args:
            corner_style (LineJoint, optional): Corner style. Defaults to ``LineJoint.ROUND``.
            cap_style (LineCap, optional): Cap style. Defaults to ``LineCap.BUTT``.

        Returns:
            None:
        """
        super().__init__()
        self.prop_corner_style = corner_style
        self.prop_cap_style = cap_style

    # region Overridden Methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.LineProperties",)
        return self._supported_services_values

    # endregion Overridden Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TCornerCaps], obj: object) -> _TCornerCaps: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TCornerCaps], obj: object, **kwargs) -> _TCornerCaps: ...

    @classmethod
    def from_obj(cls: Type[_TCornerCaps], obj: Any, **kwargs) -> _TCornerCaps:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            LineProperties: New instance.
        """
        inst = cls(**kwargs)

        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to Line Properties")

        props = {"LineCap", "LineJoint"}

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
    def prop_corner_style(self) -> LineJoint:
        """Gets/Sets the color."""
        return self._get("LineJoint")

    @prop_corner_style.setter
    def prop_corner_style(self, value: LineJoint):
        self._set("LineJoint", value)

    @property
    def prop_cap_style(self) -> LineCap:
        """Gets/Sets the color."""
        return self._get("LineCap")

    @prop_cap_style.setter
    def prop_cap_style(self, value: LineCap):
        self._set("LineCap", value)

    # endregion Properties
