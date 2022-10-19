from __future__ import annotations
from enum import Enum


class ChartTypeBase(str, Enum):
    def __str__(self) -> str:
        return str(self.value)

    def to_namespace(self) -> str:
        return f"com.sun.star.chart2.template.{self.value}"


class ColumnStackedKind(ChartTypeBase):
    COLUMN = "Column"
    STACKED_COLUMN = "StackedColumn"
    PERCENT_STACKED_COLUMN = "PercentStackedColumn"
    COLUMN_WITH_LINE = "ColumnWithLine"
    STACKED_COLUMN_WITH_LINE = "StackedColumnWithLine"


class ColumnPercentKind(ChartTypeBase):
    COLUMN_DEEP_3D = "ThreeDColumnDeep"
    COLUMN_FLAT_3D = "ThreeDColumnFlat"


class Column3dKind(ChartTypeBase):
    STACKED_3D_COLUMN_FLAT = "StackedThreeDColumnFlat"
    PERCENT_STACKED_3D_COLUMN_FLAT = "PercentStackedThreeDColumnFlat"


class BarStackedKind(ChartTypeBase):
    BAR = "Bar"
    STACKED_BAR = "StackedBar"
    PERCENT_STACKED_BAR = "PercentStackedBar"


class BarPercentKind(ChartTypeBase):
    BAR_DEEP_3D = "ThreeDBarDeep"
    BAR_FLAT_3D = "ThreeDBarFlat"


class Bar3dKind(ChartTypeBase):
    STACKED_3D_BAR_FLAT = "StackedThreeDBarFlat"
    PERCENT_STACKED_3D_BAR_FLAT = "PercentStackedThreeDBarFlat"


class PieDonutKind(ChartTypeBase):
    PIE = "Pie"
    DONUT = "Donut"


class PieExplodeKind(ChartTypeBase):
    DONUT_3D = "ThreeDDonut"
    DONUT_ALL_EXPLODED_3D = "ThreeDDonutAllExploded"


class Pie3dKind(ChartTypeBase):
    PIE_3D = "ThreeDPie"
    PIE_ALL_EXPLODED_3D = "ThreeDPieAllExploded"
    DONUT_3D = "ThreeDDonut"
    DONUT_ALL_EXPLODED_3D = "ThreeDDonutAllExploded"


class AreaStackedKind(ChartTypeBase):
    AREA = "Area"
    STACKED_AREA = "StackedArea"
    PERCENT_STACKED_AREA = "PercentStackedArea"


class AreaPercentKind(ChartTypeBase):
    AREA_3D = "ThreeDArea"
    STACKED_AREA_3D = "StackedThreeDArea"


class Area3dKind(ChartTypeBase):
    PERCENT_STACKED_AREA_3D = "PercentStackedThreeDArea"


class LineSymbolKind(ChartTypeBase):
    LINE = "Line"
    SYMBOL = "Symbol"
    LINE_SYMBOL = "LineSymbol"


class LineStackedKind(ChartTypeBase):
    STACKED_LINE = "StackedLine"
    STACKED_SYMBOL = "StackedSymbol"
    STACKED_LINE_SYMBOL = "StackedLineSymbol"


class LinePercentKind(ChartTypeBase):
    PERCENT_STACKED_LINE = "PercentStackedLine"
    PERCENT_STACKED_SYMBOL = "PercentStackedSymbol"


class Line3dKind(ChartTypeBase):
    PERCENT_STACKED_LINE_SYMBOL = "PercentStackedLineSymbol"
    LINE_3D = "ThreeDLine"
    LINE_DEEP_3D = "ThreeDLineDeep"
    STACKED_LINE_3D = "StackedThreeDLine"
    PERCENT_STACKED_LINE_3D = "PercentStackedThreeDLine"


class XYLineKind(ChartTypeBase):
    SCATTER_SYMBOL = "ScatterSymbol"
    SCATTER_LINE = "ScatterLine"
    SCATTER_LINE_SYMBOL = "ScatterLineSymbol"


class XY3dKind(ChartTypeBase):
    SCATTER_3D = "ThreeDScatter"


class BubbleKind(ChartTypeBase):
    BUBBLE = "Bubble"


class NetLineKind(ChartTypeBase):
    NET = "Net"
    NET_LINE = "NetLine"
    NET_SYMBOL = "NetSymbol"
    FILLED_NET = "FilledNet"


class NetSymbolKind(ChartTypeBase):
    STACKED_NET = "StackedNet"
    STACKED_NET_LINE = "StackedNetLine"


class NetFilledKind(ChartTypeBase):
    STACKED_NET_SYMBOL = "StackedNetSymbol"
    STACKED_FILLED_NET = "StackedFilledNet"


class NetStackedKind(ChartTypeBase):
    PERCENT_STACKED_NET = "PercentStackedNet"
    PERCENT_STACKED_NET_LINE = "PercentStackedNetLine"
    PERCENT_STACKED_NET_SYMBOL = "PercentStackedNetSymbol"


class NetPercentKind(ChartTypeBase):
    PERCENT_STACKED_FILLED_NET = "PercentStackedFilledNet"


class StockOpenKind(ChartTypeBase):
    STOCK_LOW_HIGH_CLOSE = "StockLowHighClose"


class StockVolumeKind(ChartTypeBase):
    STOCK_OPEN_LOW_HIGH_CLOSE = "StockOpenLowHighClose"
    STOCK_VOLUME_LOW_HIGH_CLOSE = "StockVolumeLowHighClose"
    STOCK_VOLUME_OPEN_LOW_HIGH_CLOSE = "StockVolumeOpenLowHighClose"


class ChartTypes:
    class Column:
        TYPE_3D = Column3dKind
        TYPE_PERCENT = ColumnPercentKind
        TYPE_STACKED = ColumnStackedKind

    class Bar:
        TYPE_3D = Bar3dKind
        TYPE_PERCENT = BarPercentKind
        TYPE_STACKED = BarStackedKind

    class Pie:
        TYPE_3D = Pie3dKind
        TYPE_DONUT = PieDonutKind
        TYPE_EXPLODE = PieExplodeKind

    class Area:
        TYPE_3D = Area3dKind
        TYPE_PERCENT = AreaPercentKind
        TYPE_STACKED = AreaStackedKind

    class Line:
        TYPE_3D = Line3dKind
        TYPE_PERCENT = LinePercentKind
        TYPE_STACKED = LineStackedKind
        TYPE_SYMBOL = LineSymbolKind

    class XY:
        TYPE_3D = XY3dKind
        TYPE_LINE = XYLineKind

    class Bubble:
        TYPE_BUBBLE = BubbleKind

    class Net:
        TYPE_FILLED = NetFilledKind
        TYPE_LINE = NetLineKind
        TYPE_PERCENT = NetPercentKind
        TYPE_STACKED = NetStackedKind
        TYPE_SYMBOL = NetSymbolKind

    class Stock:
        TYPE_OPEN = StockOpenKind
        TYPE_VOLUME = StockVolumeKind
