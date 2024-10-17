from __future__ import annotations
from random import random
from enum import Enum

from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.document import MacroExecMode

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.calc import Calc
from ooodev.office.chart import Chart, ChartDiagramKind, ChartDataCaptionEnum, Intensity
from ooodev.utils.file_io import FileIO
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.type_var import PathOrStr
from ooodev.conn.cache import Cache


class ChartKind(str, Enum):
    AREA = "area"
    BAR = "bar"
    BUBBLE = "bubble"
    DONUT = "donut"
    NET = "net"
    NET_FILLED = "net_filled"
    LINE = "line"
    PIE = "pie"
    STOCK = "stock"
    XY = "xy"


class ChartViews:
    _CHART_NAME = "chart$$_"

    def __init__(self, data_fnm: PathOrStr, chart_kind: ChartKind) -> None:
        _ = FileIO.is_exist_file(data_fnm, True)
        self._data_fnm = FileIO.get_absolute_path(data_fnm)
        self._chart_kind = chart_kind
        self._chart_name = ChartViews._CHART_NAME + str(int(random() * 10_000))

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe(), cache_obj=Cache(profile_path="", no_shared_ext=True))

        try:
            doc = Calc.open_doc(
                fnm=self._data_fnm, loader=loader, MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE_NO_WARN
            )  # type: ignore
            GUI.set_visible(visible=True, doc=doc)
            sheet = Calc.get_sheet(doc=doc, idx=0)

            if self._chart_kind == ChartKind.AREA:
                self._area_chart(doc=doc, sheet=sheet)
            elif self._chart_kind == ChartKind.BAR:
                self._bar_chart(doc=doc, sheet=sheet)
            elif self._chart_kind == ChartKind.BUBBLE:
                self._bubble_chart(doc=doc, sheet=sheet)
            elif self._chart_kind == ChartKind.DONUT:
                self._donut_chart(doc=doc, sheet=sheet)
            elif self._chart_kind == ChartKind.NET:
                self._net_chart(doc=doc, sheet=sheet)
            elif self._chart_kind == ChartKind.NET_FILLED:
                self._net_filled_chart(doc=doc, sheet=sheet)
            elif self._chart_kind == ChartKind.LINE:
                self._line_chart(doc=doc, sheet=sheet)
            elif self._chart_kind == ChartKind.PIE:
                self._pie_chart(doc=doc, sheet=sheet)
            elif self._chart_kind == ChartKind.STOCK:
                self._stock_prices_chart(doc=doc, sheet=sheet)
            elif self._chart_kind == ChartKind.XY:
                self._xy_chart(doc=doc, sheet=sheet)

            Lo.delay(2000)
            msg_result = MsgBox.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                Lo.close_doc(doc=doc, deliver_ownership=True)
                Lo.close_office()
            else:
                print("Keeping document open")
        except Exception:
            Lo.close_office()
            raise

    def _area_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        # draws an area (stacked) chart;
        # uses "Trends in Enrollment in Public Schools in the US" table
        cell_name = "A52"
        pos = Calc.get_cell_pos(sheet, cell_name)
        range_addr = Calc.get_address(sheet=sheet, range_name="E45:G50")
        chart_doc = Chart.insert_chart(
            sheet=sheet,
            chart_name=self._chart_name,
            cells_range=range_addr,
            x=round(pos.X / 1_000),
            y=round(pos.Y / 1_000),
            width=15,
            height=11,
            diagram_name=ChartDiagramKind.AREA,
        )
        Calc.goto_cell(cell_name="A43", doc=doc)

        Chart.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E43"))
        Chart.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E45"))
        Chart.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="F44"))
        Chart.view_legend(chart_doc=chart_doc, is_visible=True)

    def _bubble_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        self._chart(doc=doc, sheet=sheet, diagram_kind=ChartDiagramKind.BUBBLE)

    def _bar_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        self._chart(doc=doc, sheet=sheet, diagram_kind=ChartDiagramKind.BAR)

    def _chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet, diagram_kind: ChartDiagramKind) -> None:
        # uses "Sneakers Sold this Month" table
        range_addr = Calc.get_address(sheet=sheet, range_name="A2:B8")
        chart_doc = Chart.insert_chart(
            sheet=sheet,
            chart_name=self._chart_name,
            cells_range=range_addr,
            width=15,
            height=11,
            diagram_name=diagram_kind,
        )
        Calc.goto_cell(cell_name="A1", doc=doc)

        Chart.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A1"))
        Chart.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A2"))
        Chart.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B2"))

    def _donut_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        # draws a 3D donut chart with 2 rings
        # uses the "Annual Expenditure on Institutions" table
        cell_name = "D43"
        range_addr = Calc.get_address(sheet=sheet, range_name="A44:C50")
        pos = Calc.get_cell_pos(sheet, cell_name)
        chart_doc = Chart.insert_chart(
            sheet=sheet,
            chart_name=self._chart_name,
            cells_range=range_addr,
            x=round(pos.X / 1_000),
            y=round(pos.Y / 1_000),
            width=15,
            height=11,
            diagram_name=ChartDiagramKind.DONUT,
        )
        Calc.goto_cell(cell_name="A44", doc=doc)

        Chart.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A43"))
        Chart.view_legend(chart_doc=chart_doc, is_visible=True)
        subtitle = f'Outer: {Calc.get_string(sheet=sheet, cell_name="B44")}\nInner: {Calc.get_string(sheet=sheet, cell_name="C44")}'
        Chart.set_subtitle(chart_doc=chart_doc, subtitle=subtitle)

    def _net_filled_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        # draws a net chart;
        # uses the "No of Calls per Day" table
        cell_name = "E55"
        range_addr = Calc.get_address(sheet=sheet, range_name="A56:D63")
        pos = Calc.get_cell_pos(sheet, cell_name)
        chart_doc = Chart.insert_chart(
            sheet=sheet,
            chart_name=self._chart_name,
            cells_range=range_addr,
            x=round(pos.X / 1_000),
            y=round(pos.Y / 1_000),
            width=15,
            height=11,
            diagram_name=ChartDiagramKind.FILLED_NET,
        )
        Calc.goto_cell(cell_name=cell_name, doc=doc)

        Chart.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A55"))
        Chart.view_legend(chart_doc=chart_doc, is_visible=True)
        Chart.set_data_caption(chart_doc=chart_doc, label_types=ChartDataCaptionEnum.NONE)

    def _net_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        # draws a net chart;
        # uses the "No of Calls per Day" table
        cell_name = "E55"
        range_addr = Calc.get_address(sheet=sheet, range_name="A56:D63")
        pos = Calc.get_cell_pos(sheet, cell_name)
        chart_doc = Chart.insert_chart(
            sheet=sheet,
            chart_name=self._chart_name,
            cells_range=range_addr,
            x=round(pos.X / 1_000),
            y=round(pos.Y / 1_000),
            width=15,
            height=11,
            diagram_name=ChartDiagramKind.NET,
        )
        Calc.goto_cell(cell_name=cell_name, doc=doc)

        Chart.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A55"))
        Chart.view_legend(chart_doc=chart_doc, is_visible=True)
        Chart.set_data_caption(chart_doc=chart_doc, label_types=ChartDataCaptionEnum.NONE)

    def _line_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        # draw a line chart with data points, no legend;
        # uses "Humidity Levels in NY" table
        range_addr = Calc.get_address(sheet=sheet, range_name="A14:B21")
        chart_doc = Chart.insert_chart(
            sheet=sheet,
            chart_name=self._chart_name,
            cells_range=range_addr,
            width=15,
            height=11,
            diagram_name=ChartDiagramKind.LINE,
        )
        Calc.goto_cell(cell_name="A1", doc=doc)

        Chart.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A13"))
        Chart.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A14"))
        Chart.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B14"))
        Chart.set_area_transparency(chart_doc=chart_doc, val=Intensity(30))

    def _pie_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        # draw a pie chart, with legend and subtitle;
        # uses "Top 5 States with the Most Elementary and Secondary Schools"
        range_addr = Calc.get_address(sheet=sheet, range_name="E2:F8")
        chart_doc = Chart.insert_chart(
            sheet=sheet,
            chart_name=self._chart_name,
            cells_range=range_addr,
            width=15,
            height=11,
            diagram_name=ChartDiagramKind.PIE,
        )
        Calc.goto_cell(cell_name="A1", doc=doc)

        Chart.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E1"))
        Chart.set_subtitle(chart_doc=chart_doc, subtitle=Calc.get_string(sheet=sheet, cell_name="F2"))
        Chart.view_legend(chart_doc=chart_doc, is_visible=True)

    def _stock_prices_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        # draws a stock chart, with an extra pork bellies line
        cell_name = "E148"
        range_addr = Calc.get_address(sheet=sheet, range_name="E141:I146")
        pos = Calc.get_cell_pos(sheet, cell_name)
        chart_doc = Chart.insert_chart(
            sheet=sheet,
            chart_name=self._chart_name,
            cells_range=range_addr,
            x=round(pos.X / 1_000),
            y=round(pos.Y / 1_000),
            width=15,
            height=11,
            diagram_name=ChartDiagramKind.STOCK,
        )
        Calc.goto_cell(cell_name="A148", doc=doc)

        Chart.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E140"))
        Chart.set_data_caption(chart_doc=chart_doc, label_types=ChartDataCaptionEnum.NONE)
        Chart.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E141"))
        Chart.set_y_axis_title(chart_doc=chart_doc, title="Dollars")

        Chart.view_legend(chart_doc=chart_doc, is_visible=True)

    def _xy_chart(self, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        # data from http://www.mathsisfun.com/data/scatter-xy-plots.html;
        # uses the "Ice Cream Sales vs Temperature" table
        cell_name = "E148"
        range_addr = Calc.get_address(sheet=sheet, range_name="A110:B122")
        pos = Calc.get_cell_pos(sheet, cell_name)
        chart_doc = Chart.insert_chart(
            sheet=sheet,
            chart_name=self._chart_name,
            cells_range=range_addr,
            x=round(pos.X / 1_000),
            y=round(pos.Y / 1_000),
            width=16,
            height=11,
            diagram_name=ChartDiagramKind.XY,
        )
        Calc.goto_cell(cell_name="A148", doc=doc)

        Chart.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A109"))
        Chart.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A110"))
        Chart.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B110"))
