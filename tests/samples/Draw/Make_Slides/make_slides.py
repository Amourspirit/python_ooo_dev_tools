from __future__ import annotations

import uno
from com.sun.star.drawing import XDrawPage
from com.sun.star.lang import XComponent
from com.sun.star.drawing import XShape

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.loader.lo import Lo
from ooodev.office.draw import Draw, DrawingShapeKind, LineStyle
from ooodev.gui.gui import GUI
from ooodev.utils.props import Props
from ooodev.utils.gallery import Gallery


class MakeSlides:
    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        doc = Draw.create_draw_doc(loader)

        try:
            GUI.set_visible(is_visible=True, odoc=doc)
            Lo.delay(1_000)  # need delay or zoom may not occur
            GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

            curr_slide = Draw.get_slide(doc=doc, idx=0)
            shape_type = DrawingShapeKind.GRAPHIC_OBJECT_SHAPE  # DrawingShapeKind.CUSTOM_SHAPE
            smile_shape = Draw.make_shape(shape_type=shape_type, x=10, y=10, height=50, width=50)
            g_name = "Flower"  # "Callout-2" "Diamond-Bevel" "Sun"
            graphic = Gallery.find_gallery_graphic(g_name)
            Draw.set_image_graphic(shape=smile_shape, graphic=graphic)
            Draw.set_line_style(shape=smile_shape, style=LineStyle.NONE)
            curr_slide.add(smile_shape)

            Lo.delay(2000)
            msg_result = MsgBox.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                Lo.close_doc(doc=doc)
            else:
                print("Keeping document open")
        except Exception:
            Lo.close_doc(doc=doc)
            raise

    def _add_dispatch_shape(
        self, slide: XDrawPage, shape_dispatch: str, x: int, y: int, width: int, height: int
    ) -> XShape:
        pass

    def _create_dispatch_shape(self, slide: XDrawPage, shape_dispatch: str) -> XShape:
        pass

    def _dispatch_shapes(self, doc: XComponent) -> None:
        curr_slide = Draw.add_slide(doc)
        Draw.title_only_slide(slide=curr_slide, header="Dispatched Shapes")

        GUI.set_visible(is_visible=True, odoc=doc)
        Lo.delay(1000)

        Draw.goto_page(doc=doc, page=curr_slide)
        print(f"Viewing Slide number: {Draw.get_slide_number(Draw.get_viewed_page(doc))}")

        # first row
        d_shape = Draw.create_control_shape()
