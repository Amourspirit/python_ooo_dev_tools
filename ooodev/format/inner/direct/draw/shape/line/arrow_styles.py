from __future__ import annotations
from typing import Tuple, overload, Any, TYPE_CHECKING, TypeVar, Type
import uno

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import props as mProps
from ooodev.utils.kind.graphic_arrow_style_kind import GraphicArrowStyleKind

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TArrowStyles = TypeVar(name="_TArrowStyles", bound="ArrowStyles")


class ArrowStyles(StyleBase):
    """
    This class represents the line Arrow styles of a shape or line.
    """

    def __init__(
        self,
        start_line_name: GraphicArrowStyleKind | str | None = None,
        start_line_width: float | UnitT | None = None,
        start_line_center: bool | None = None,
        end_line_name: GraphicArrowStyleKind | str | None = None,
        end_line_width: float | UnitT | None = None,
        end_line_center: bool | None = None,
    ) -> None:
        """
        Constructor.

        Args:
            start_line_name (GraphicArrowStyleKind, str, optional): Start line name. Defaults to ``None``.
            start_line_width (float, UnitT, optional): Start line width in mm units. Defaults to ``None``.
            start_line_center (bool, optional): Start line center. Defaults to ``None``.
            end_line_name (GraphicArrowStyleKind, str, optional): End line name. Defaults to ``None``.
            end_line_width (float, UnitT, optional): End line width in mm units. Defaults to ``None``.
            end_line_center (bool, optional): End line center. Defaults to ``None``.

        Returns:
            None:
        """
        super().__init__()
        self.prop_start_line_name = start_line_name
        self.prop_start_line_width = start_line_width
        self.prop_start_line_center = start_line_center
        self.prop_end_line_name = end_line_name
        self.prop_end_line_width = end_line_width
        self.prop_end_line_center = end_line_center

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
    def from_obj(cls: Type[_TArrowStyles], obj: object) -> _TArrowStyles: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TArrowStyles], obj: object, **kwargs) -> _TArrowStyles: ...

    @classmethod
    def from_obj(cls: Type[_TArrowStyles], obj: Any, **kwargs) -> _TArrowStyles:
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

        props = {"LineStartName", "LineStartWidth", "LineStartCenter", "LineEndName", "LineEndWidth", "LineEndCenter"}

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
    def prop_start_line_name(self) -> str | None:
        """Gets/Sets start line name."""
        return self._get("LineStartName")

    @prop_start_line_name.setter
    def prop_start_line_name(self, value: GraphicArrowStyleKind | str | None):
        if value is None:
            self._remove("LineStartName")
            return
        self._set("LineStartName", str(value))

    @property
    def prop_start_line_width(self) -> UnitMM | None:
        """Gets/Sets the start line width."""
        pv = self._get("LineStartWidth")
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_start_line_width.setter
    def prop_start_line_width(self, value: UnitT | float | None):
        if value is None:
            self._remove("LineStartWidth")
            return
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            val = UnitMM(value).get_value_mm100()  # type: ignore
        self._set("LineStartWidth", val)

    @property
    def prop_start_line_center(self) -> bool | None:
        """Gets/Sets the start line center."""
        return self._get("LineStartCenter")

    @prop_start_line_center.setter
    def prop_start_line_center(self, value: bool | None):
        if value is None:
            self._remove("LineStartCenter")
            return
        self._set("LineStartCenter", value)

    @property
    def prop_end_line_name(self) -> str | None:
        """Gets/Sets end line name."""
        return self._get("LineEndName")

    @prop_end_line_name.setter
    def prop_end_line_name(self, value: GraphicArrowStyleKind | str | None):
        if value is None:
            self._remove("LineEndName")
            return
        self._set("LineEndName", str(value))

    @property
    def prop_end_line_width(self) -> UnitMM | None:
        """Gets/Sets the end line width."""
        pv = self._get("LineEndWidth")
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_end_line_width.setter
    def prop_end_line_width(self, value: UnitT | float | None):
        if value is None:
            self._remove("LineEndWidth")
            return
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            val = UnitMM(value).get_value_mm100()  # type: ignore
        self._set("LineEndWidth", val)

    @property
    def prop_end_line_center(self) -> bool | None:
        """Gets/Sets the end line center."""
        return self._get("LineEndCenter")

    @prop_end_line_center.setter
    def prop_end_line_center(self, value: bool | None):
        if value is None:
            self._remove("LineEndCenter")
            return
        self._set("LineEndCenter", value)

    # endregion Properties
