from __future__ import annotations
from pathlib import Path
import tempfile

from com.sun.star.frame import XComponentLoader

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.calc import Calc
from ooodev.office.chart2 import Chart2, Angle
from ooodev.office.draw import Draw, DrawingNameSpaceKind, mEx
from ooodev.utils.dispatch.global_edit_dispatch import GlobalEditDispatch
from ooodev.utils.file_io import FileIO
from ooodev.gui.gui import GUI
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.images_lo import ImagesLo
from ooodev.loader.lo import Lo
from ooodev.utils.type_var import PathOrStr


class SlideChart:
    def __init__(self, data_fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(data_fnm, True)
        self._data_fnm = FileIO.get_absolute_path(data_fnm)
        self._out_dir = Path(tempfile.mkdtemp())

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            has_chart = self._make_col_chart(loader)

            doc = Draw.create_impress_doc(loader)
            # to make the construction visible
            GUI.set_visible(is_visible=True, odoc=doc)

            # access first page.
            slide = Draw.get_slide(doc=doc, idx=0)
            body = Draw.bullets_slide(slide=slide, title="Sneakers Are Selling!")
            Draw.add_bullet(bulls_txt=body, level=0, text="Sneaker profits have increased")

            if has_chart:
                Lo.delay(1_000)
                Lo.dispatch_cmd(GlobalEditDispatch.PASTE)

            try:
                ole_shape = Draw.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.OLE2_SHAPE)
                slide_size = Draw.get_slide_size(slide)
                shape_size = Draw.get_size(ole_shape)
                shape_pos = Draw.get_position(ole_shape)

                y = slide_size.height - shape_size.height - 20
                # move pic down
                Draw.set_position(shape=ole_shape, x=shape_pos.X, y=y)
            except mEx.ShapeMissingError:
                Lo.print("Did not find shape, unable to set size and position")

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

    def _make_col_chart(self, loader: XComponentLoader) -> bool:
        ssdoc = Calc.open_doc(fnm=self._data_fnm, loader=loader)
        try:
            GUI.set_visible(is_visible=True, odoc=ssdoc)  # or selection is not copied
            sheet = Calc.get_sheet(doc=ssdoc, index=0)

            range_addr = Calc.get_address(sheet=sheet, range_name="A2:B8")
            chart_doc = Chart2.insert_chart(
                sheet=sheet,
                cells_range=range_addr,
                cell_name="C3",
                width=10,
                height=8,
                diagram_name=ChartTypes.Column.TEMPLATE_STACKED.COLUMN,
            )

            Chart2.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A1"))
            Chart2.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A2"))
            Chart2.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B2"))
            Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
            Lo.delay(1_000)
            Chart2.copy_chart(ssdoc=ssdoc, sheet=sheet)

            try:
                ImagesLo.save_graphic(pic=Chart2.get_chart_image(sheet), fnm=Path(self._out_dir, "chartImage.png"))
            except mEx.ImageError:
                pass

            return True
        except Exception as e:
            Lo.print("Error making col chart")
            Lo.print(f"  {e}")
        finally:
            Lo.close_doc(doc=ssdoc)
        return False
