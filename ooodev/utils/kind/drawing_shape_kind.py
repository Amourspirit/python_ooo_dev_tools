from .kind_base import KindBase
from ooodev.utils.kind import kind_helper


class DrawingShapeKind(KindBase):
    """
    Values used with ``com.sun.star.drawing.*``

    See Also:
        - `Drawing API <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html>`_
        - `OpenOffice Shape Types <https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Shape_Types>`_
    """

    APPLET_SHAPE = "AppletShape"
    CAPTION_SHAPE = "CaptionShape"
    CHART_LEGEND = "ChartLegend"
    CHART_TITLE = "ChartTitle"
    CLOSED_BEZIER_SHAPE = "ClosedBezierShape"
    CONNECTOR_SHAPE = "ConnectorShape"
    CONTROL_SHAPE = "ControlShape"
    CUSTOM_SHAPE = "CustomShape"
    ELLIPSE_SHAPE = "EllipseShape"
    GRAPHIC_OBJECT_SHAPE = "GraphicObjectShape"
    GROUP_SHAPE = "GroupShape"
    LINE_SHAPE = "LineShape"
    MEASURE_SHAPE = "MeasureShape"
    MEDIA_SHAPE = "MediaShape"
    OLE2_SHAPE = "OLE2Shape"
    OPEN_BEZIER_SHAPE = "OpenBezierShape"
    PAGE_SHAPE = "PageShape"
    PLUGIN_SHAPE = "PluginShape"
    POLY_LINE_SHAPE = "PolyLineShape"
    POLY_POLYGON_BEZIER_SHAPE = "PolyPolygonBezierShape"
    POLY_POLYGON_SHAPE = "PolyPolygonShape"
    RECTANGLE_SHAPE = "RectangleShape"
    TEXT_SHAPE = "TextShape"

    # could not find MediaShape in api.
    # https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html
    # however it can be found in examples.
    # https://ask.libreoffice.org/t/how-to-add-video-to-impress-with-python/33050/2

    def to_namespace(self) -> str:
        """Gets full name-space value of instance"""
        return f"com.sun.star.drawing.{self.value}"

    @staticmethod
    def from_str(s: str) -> "DrawingShapeKind":
        """
        Gets an ``DrawingShapeKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DrawingShapeKind`` instance.

        Returns:
            DrawingShapeKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, DrawingShapeKind)
