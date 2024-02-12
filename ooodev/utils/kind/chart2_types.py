from __future__ import annotations
from enum import Enum


class ChartBaseTypeEnum(str, Enum):
    def __str__(self) -> str:
        return str(self.value)

    def to_namespace(self) -> str:
        raise NotImplementedError


class ChartTemplateBase(ChartBaseTypeEnum):
    """
    Base Enum for all enums found in :py:class:`~.chart2_types.ChartTypes`

    Example:

        :py:attr:`.Chart2.ChartLookup` is an alias of :py:class:`~.chart2_types.ChartTypes`

        .. code-block:: python

            Chart2.has_categories(diagram_name=Chart2.ChartLookup.Bar.TYPE_PERCENT.BAR_DEEP_3D)

    See Also:
        :py:meth:`.Chart2.has_categories`
    """

    def to_namespace(self) -> str:
        return f"com.sun.star.chart2.template.{self.value}"


class ChartTypeNameBase(ChartBaseTypeEnum):
    def to_namespace(self) -> str:
        return f"com.sun.star.chart2.{self.value}"


class ColumnStackedKind(ChartTemplateBase):
    COLUMN = "Column"
    STACKED_COLUMN = "StackedColumn"
    PERCENT_STACKED_COLUMN = "PercentStackedColumn"


class ColumnPercentKind(ChartTemplateBase):
    COLUMN_DEEP_3D = "ThreeDColumnDeep"
    COLUMN_FLAT_3D = "ThreeDColumnFlat"


class Column3dKind(ChartTemplateBase):
    STACKED_3D_COLUMN_FLAT = "StackedThreeDColumnFlat"
    PERCENT_STACKED_3D_COLUMN_FLAT = "PercentStackedThreeDColumnFlat"


class BarStackedKind(ChartTemplateBase):
    BAR = "Bar"
    STACKED_BAR = "StackedBar"
    PERCENT_STACKED_BAR = "PercentStackedBar"


class BarPercentKind(ChartTemplateBase):
    BAR_DEEP_3D = "ThreeDBarDeep"
    BAR_FLAT_3D = "ThreeDBarFlat"


class Bar3dKind(ChartTemplateBase):
    STACKED_3D_BAR_FLAT = "StackedThreeDBarFlat"
    PERCENT_STACKED_3D_BAR_FLAT = "PercentStackedThreeDBarFlat"


class PieDonutKind(ChartTemplateBase):
    PIE = "Pie"
    DONUT = "Donut"


class PieExplodeKind(ChartTemplateBase):
    PIE_ALL_EXPLODED = "PieAllExploded"
    DONUT_ALL_EXPLODED = "DonutAllExploded"


class Pie3dKind(ChartTemplateBase):
    PIE_3D = "ThreeDPie"
    PIE_ALL_EXPLODED_3D = "ThreeDPieAllExploded"
    DONUT_3D = "ThreeDDonut"
    DONUT_ALL_EXPLODED_3D = "ThreeDDonutAllExploded"


class AreaStackedKind(ChartTemplateBase):
    AREA = "Area"
    STACKED_AREA = "StackedArea"
    PERCENT_STACKED_AREA = "PercentStackedArea"


class AreaPercentKind(ChartTemplateBase):
    AREA_3D = "ThreeDArea"
    STACKED_AREA_3D = "StackedThreeDArea"


class Area3dKind(ChartTemplateBase):
    PERCENT_STACKED_AREA_3D = "PercentStackedThreeDArea"


class LineSymbolKind(ChartTemplateBase):
    LINE = "Line"
    SYMBOL = "Symbol"
    LINE_SYMBOL = "LineSymbol"


class LineStackedKind(ChartTemplateBase):
    STACKED_LINE = "StackedLine"
    STACKED_SYMBOL = "StackedSymbol"
    STACKED_LINE_SYMBOL = "StackedLineSymbol"


class LinePercentKind(ChartTemplateBase):
    PERCENT_STACKED_LINE = "PercentStackedLine"
    PERCENT_STACKED_SYMBOL = "PercentStackedSymbol"


class Line3dKind(ChartTemplateBase):
    PERCENT_STACKED_LINE_SYMBOL = "PercentStackedLineSymbol"
    LINE_3D = "ThreeDLine"
    LINE_DEEP_3D = "ThreeDLineDeep"
    STACKED_LINE_3D = "StackedThreeDLine"
    PERCENT_STACKED_LINE_3D = "PercentStackedThreeDLine"


class XYLineKind(ChartTemplateBase):
    SCATTER_SYMBOL = "ScatterSymbol"
    SCATTER_LINE = "ScatterLine"
    SCATTER_LINE_SYMBOL = "ScatterLineSymbol"


class XY3dKind(ChartTemplateBase):
    SCATTER_3D = "ThreeDScatter"


class BubbleKind(ChartTemplateBase):
    BUBBLE = "Bubble"


class NetLineKind(ChartTemplateBase):
    NET = "Net"
    NET_LINE = "NetLine"
    NET_SYMBOL = "NetSymbol"
    FILLED_NET = "FilledNet"


class NetSymbolKind(ChartTemplateBase):
    STACKED_NET = "StackedNet"
    STACKED_NET_LINE = "StackedNetLine"


class NetFilledKind(ChartTemplateBase):
    STACKED_NET_SYMBOL = "StackedNetSymbol"
    STACKED_FILLED_NET = "StackedFilledNet"


class NetStackedKind(ChartTemplateBase):
    PERCENT_STACKED_NET = "PercentStackedNet"
    PERCENT_STACKED_NET_LINE = "PercentStackedNetLine"
    PERCENT_STACKED_NET_SYMBOL = "PercentStackedNetSymbol"


class NetPercentKind(ChartTemplateBase):
    PERCENT_STACKED_FILLED_NET = "PercentStackedFilledNet"


class StockOpenKind(ChartTemplateBase):
    STOCK_LOW_HIGH_CLOSE = "StockLowHighClose"


class StockVolumeKind(ChartTemplateBase):
    STOCK_OPEN_LOW_HIGH_CLOSE = "StockOpenLowHighClose"
    STOCK_VOLUME_LOW_HIGH_CLOSE = "StockVolumeLowHighClose"
    STOCK_VOLUME_OPEN_LOW_HIGH_CLOSE = "StockVolumeOpenLowHighClose"


class ColumnAndLineStackedKind(ChartTemplateBase):
    COLUMN_WITH_LINE = "ColumnWithLine"
    STACKED_COLUMN_WITH_LINE = "StackedColumnWithLine"


class NamedColumnKind(ChartTypeNameBase):
    COLUMN_CHART = "ColumnChartType"


class NamedBarKind(ChartTypeNameBase):
    BAR_CHART = "BarChartType"


class NamedPieKind(ChartTypeNameBase):
    PIE_CHART = "PieChartType"


class NamedAreaKind(ChartTypeNameBase):
    AREA_CHART = "AreaChartType"


class NamedLineKind(ChartTypeNameBase):
    LINE_CHART = "LineChartType"


class NamedXYKind(ChartTypeNameBase):
    SCATTER_CHART = "ScatterChartType"


class NamedBubbleKind(ChartTypeNameBase):
    BUBBLE_CHART = "BubbleChartType"


class NamedNetKind(ChartTypeNameBase):
    NET_CHART = "NetChartType"
    FILLED_NET_CHART = "FilledNetChartType"


class NamedStockKind(ChartTypeNameBase):
    CANDLE_STICK_CHART = "CandleStickChartType"


class ChartTypes:
    """
    Class for convenient lookup of chart type names.
    """

    class Column:
        DEFAULT = ColumnStackedKind.COLUMN
        NAMED = NamedColumnKind
        TEMPLATE_3D = Column3dKind
        TEMPLATE_PERCENT = ColumnPercentKind
        TEMPLATE_STACKED = ColumnStackedKind

    class ColumnAndLine:
        DEFAULT = ColumnAndLineStackedKind.COLUMN_WITH_LINE
        TEMPLATE_STACKED = ColumnAndLineStackedKind

    class Bar:
        DEFAULT = BarStackedKind.BAR
        NAMED = NamedBarKind
        TEMPLATE_3D = Bar3dKind
        TEMPLATE_PERCENT = BarPercentKind
        TEMPLATE_STACKED = BarStackedKind

    class Pie:
        DEFAULT = PieDonutKind.PIE
        NAMED = NamedPieKind
        TEMPLATE_3D = Pie3dKind
        TEMPLATE_DONUT = PieDonutKind
        TEMPLATE_EXPLODE = PieExplodeKind

    class Area:
        DEFAULT = AreaStackedKind.AREA
        NAMED = NamedAreaKind
        TEMPLATE_3D = Area3dKind
        TEMPLATE_PERCENT = AreaPercentKind
        TEMPLATE_STACKED = AreaStackedKind

    class Line:
        DEFAULT = LineSymbolKind.LINE
        NAMED = NamedLineKind
        TEMPLATE_3D = Line3dKind
        TEMPLATE_PERCENT = LinePercentKind
        TEMPLATE_STACKED = LineStackedKind
        TEMPLATE_SYMBOL = LineSymbolKind

    class XY:
        DEFAULT = XYLineKind.SCATTER_SYMBOL
        NAMED = NamedXYKind
        TEMPLATE_3D = XY3dKind
        TEMPLATE_LINE = XYLineKind

    class Bubble:
        DEFAULT = NamedBubbleKind.BUBBLE_CHART
        NAMED = NamedBubbleKind
        TEMPLATE_BUBBLE = BubbleKind

    class Net:
        DEFAULT = NetLineKind.NET_SYMBOL
        NAMED = NamedNetKind
        TEMPLATE_FILLED = NetFilledKind
        TEMPLATE_LINE = NetLineKind
        TEMPLATE_PERCENT = NetPercentKind
        TEMPLATE_STACKED = NetStackedKind
        TEMPLATE_SYMBOL = NetSymbolKind

    class Stock:
        DEFAULT = StockOpenKind.STOCK_LOW_HIGH_CLOSE
        NAMED = NamedStockKind
        TEMPLATE_OPEN = StockOpenKind
        TEMPLATE_VOLUME = StockVolumeKind
