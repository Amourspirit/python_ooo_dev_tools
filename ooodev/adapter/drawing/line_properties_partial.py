from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno

from ooodev.adapter.drawing.line_dash_struct_comp import LineDashStructComp
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.events.events import Events
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from com.sun.star.drawing import LineDash  # Struct
    from com.sun.star.drawing import LineProperties
    from com.sun.star.drawing import PolyPolygonBezierCoords  # Struct
    from com.sun.star.util import Color  # type def
    from ooo.dyn.drawing.line_cap import LineCap
    from ooo.dyn.drawing.line_joint import LineJoint
    from ooo.dyn.drawing.line_style import LineStyle
    from ooodev.units.unit_obj import UnitT
    from ooodev.events.args.key_val_args import KeyValArgs


class LinePropertiesPartial:
    """
    Partial class for LineProperties Service.

    See Also:
        `API LineProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineProperties.html>`_
    """

    def __init__(self, component: LineProperties) -> None:
        """
        Constructor

        Args:
            component (LineProperties): UNO Component that implements ``com.sun.star.drawing.LineProperties`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``LineProperties``.
        """
        self.__component = component
        self.__event_provider = Events(self)
        self.__props = {}

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self.__component, prop_name):
                setattr(self.__component, prop_name, event_args.source.component)

        self.__fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member
        self.__event_provider.subscribe_event(
            "com_sun_star_drawing_LineDash_changed", self.__fn_on_comp_struct_changed
        )

    # region LineProperties
    @property
    def line_cap(self) -> LineCap | None:
        """
        Gets/Sets the rendering of ends of thick lines.

        **optional**:

        Returns:
            LineCap: The line cap.

        Hint:
            - ``LineCap`` can be imported from ``ooo.dyn.drawing.line_cap``
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineCap  # type: ignore
        return None

    @line_cap.setter
    def line_cap(self, value: LineCap) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineCap = value  # type: ignore

    @property
    def line_color(self) -> Color:
        """
        Gets/Sets the line color.
        """
        return self.__component.LineColor

    @line_color.setter
    def line_color(self, value: Color) -> None:
        self.__component.LineColor = value

    @property
    def line_dash(self) -> LineDashStructComp:
        """
        Gets/Sets the dash of the line.

        When setting the value can be a ``LineDash`` or a ``LineDashStructComp``.

        Returns:
            LineDashStructComp: The line dash.

        Hint:
            - ``LineDash`` can be imported from ``ooo.dyn.drawing.line_dash``
        """
        key = "LineDash"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = LineDashStructComp(self.__component.LineDash, key, self.__event_provider)
            self.__props[key] = prop
        return cast(LineDashStructComp, prop)

    @line_dash.setter
    def line_dash(self, value: LineDash | LineDashStructComp) -> None:
        key = "LineDash"
        if mInfo.Info.is_instance(value, LineDashStructComp):
            self.__component.LineDash = value.copy()
        else:
            self.__component.LineDash = cast("LineDash", value)
        if key in self.__props:
            del self.__props[key]

    @property
    def line_dash_name(self) -> str:
        """
        Gets/Sets the name of the dash of the line.
        """
        return self.__component.LineDashName

    @line_dash_name.setter
    def line_dash_name(self, value: str) -> None:
        self.__component.LineDashName = value

    @property
    def line_end(self) -> PolyPolygonBezierCoords | None:
        """
        Gets/Sets the line end in the form of a poly polygon Bezier.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineEnd
        return None

    @line_end.setter
    def line_end(self, value: PolyPolygonBezierCoords) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineEnd = value

    @property
    def line_end_center(self) -> bool | None:
        """
        Gets/Sets the line end center.

        If this property is ``True``, the line will end in the center of the polygon.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineEndCenter
        return None

    @line_end_center.setter
    def line_end_center(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineEndCenter = value

    @property
    def line_end_name(self) -> str | None:
        """
        Gets/Sets the name of the line end poly polygon Bezier.

        If this string is empty, no line end polygon is rendered.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineEndName
        return None

    @line_end_name.setter
    def line_end_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineEndName = value

    @property
    def line_end_width(self) -> int | None:
        """
        Gets/Sets the width of the line end polygon.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineEndWidth
        return None

    @line_end_width.setter
    def line_end_width(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineEndWidth = value

    @property
    def line_joint(self) -> LineJoint:
        """
        Gets/Sets the rendering of joints between thick lines.

        Returns:
            LineJoint: The line joint.

        Hint:
            - ``LineJoint`` can be imported from ``ooo.dyn.drawing.line_joint``
        """
        return self.__component.LineJoint  # type: ignore

    @line_joint.setter
    def line_joint(self, value: LineJoint) -> None:
        self.__component.LineJoint = value  # type: ignore

    @property
    def line_start(self) -> PolyPolygonBezierCoords | None:
        """
        Gets/Sets the line start in the form of a poly polygon Bezier.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineStart
        return None

    @line_start.setter
    def line_start(self, value: PolyPolygonBezierCoords) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineStart = value

    @property
    def line_start_center(self) -> bool | None:
        """
        Gets/Sets the line start center.

        If this property is ``True``, the line will start from the center of the polygon.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineStartCenter
        return None

    @line_start_center.setter
    def line_start_center(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineStartCenter = value

    @property
    def line_start_name(self) -> str | None:
        """
        Gets/Sets the name of the line start poly polygon Bezier.

        If this string is empty, no line start polygon is rendered.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineStartName
        return None

    @line_start_name.setter
    def line_start_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineStartName = value

    @property
    def line_start_width(self) -> int | None:
        """
        Gets/Sets the width of the line start polygon.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineStartWidth
        return None

    @line_start_width.setter
    def line_start_width(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineStartWidth = value

    @property
    def line_style(self) -> LineStyle:
        """
        Gets/Sets the type of the line.

        Return:
            LineStyle: Line Style.

        Hint:
            - ``LineStyle`` can be imported from ``ooo.dyn.drawing.line_style``
        """
        return self.__component.LineStyle  # type: ignore

    @line_style.setter
    def line_style(self, value: LineStyle) -> None:
        self.__component.LineStyle = value  # type: ignore

    @property
    def line_transparence(self) -> int:
        """
        Gets/Sets the extent of transparency.
        """
        return self.__component.LineTransparence

    @line_transparence.setter
    def line_transparence(self, value: int) -> None:
        self.__component.LineTransparence = value

    @property
    def line_width(self) -> UnitMM100:
        """
        Gets/Sets the width of the line in ``1/100th mm`` units.
        """
        return UnitMM100(self.__component.LineWidth)

    @line_width.setter
    def line_width(self, value: float | UnitT) -> None:
        self.__component.LineWidth = UnitMM100.from_unit_val(value).value

    # endregion LineProperties
