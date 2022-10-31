from enum import Enum


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
