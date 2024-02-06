from __future__ import annotations
from typing import TYPE_CHECKING
import contextlib
import uno

from ooodev.units import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.drawing import LineDash  # Struct
    from com.sun.star.drawing import LineProperties
    from com.sun.star.drawing import PolyPolygonBezierCoords  # Struct
    from com.sun.star.drawing.LineCap import LineCapProto  # type: ignore
    from com.sun.star.drawing.LineJoint import LineJointProto  # type: ignore
    from com.sun.star.drawing.LineStyle import LineStyleProto  # type: ignore
    from com.sun.star.util import Color  # type def
    from ooodev.units import UnitT


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

    # region LineProperties
    @property
    def line_cap(self) -> LineCapProto | None:
        """
        Gets/Sets the rendering of ends of thick lines.

        **optional**:
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LineCap
        return None

    @line_cap.setter
    def line_cap(self, value: LineCapProto) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LineCap = value

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
    def line_dash(self) -> LineDash:
        """
        Gets/Sets the dash of the line.
        """
        return self.__component.LineDash

    @line_dash.setter
    def line_dash(self, value: LineDash) -> None:
        self.__component.LineDash = value

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
    def line_joint(self) -> LineJointProto:
        """
        Gets/Sets the rendering of joints between thick lines.
        """
        return self.__component.LineJoint

    @line_joint.setter
    def line_joint(self, value: LineJointProto) -> None:
        self.__component.LineJoint = value

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
    def line_style(self) -> LineStyleProto:
        """
        Gets/Sets the type of the line.
        """
        return self.__component.LineStyle

    @line_style.setter
    def line_style(self, value: LineStyleProto) -> None:
        self.__component.LineStyle = value

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
