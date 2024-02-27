from enum import Enum
from ooodev.utils.kind import kind_helper


class ChartDiagramKind(str, Enum):
    AREA = "AreaDiagram"
    BAR = "BarDiagram"
    BUBBLE = "BubbleDiagram"
    DONUT = "DonutDiagram"
    FILLED_NET = "FilledNetDiagram"
    LINE = "LineDiagram"
    NET = "NetDiagram"
    PIE = "PieDiagram"
    STOCK = "StockDiagram"
    XY = "XYDiagram"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "ChartDiagramKind":
        """
        Gets an ``ChartDiagramKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``DrawingHatchingKind`` instance.

        Returns:
            ChartDiagramKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, ChartDiagramKind)
