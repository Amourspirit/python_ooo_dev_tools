# - list the shapes on the default master page.
#    - add shape and text to the master page
#    - set the footer text
#    - have normal slides use the slide number and footer on the master page
#
#    - create a second master page
#    - link one of the slides to the second master page
from __future__ import annotations

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.loader.lo import Lo
from ooodev.office.draw import Draw
from ooodev.gui.gui import GUI
from ooodev.utils.props import Props
from ooodev.utils.color import CommonColor


class MasterUse:
    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())
        try:
            doc = Draw.create_impress_doc(loader)

            # report on the shapes on the default master page
            master_page = Draw.get_master_page(doc=doc, idx=0)
            print("Default Master Page")
            Draw.show_shapes_info(master_page)

            # set the master page's footer text
            Draw.set_master_footer(master=master_page, text="Master Use Slides")

            # add a rectangle and text to the default master page
            # at the top-left of the slide
            sz = Draw.get_slide_size(master_page)
            _ = Draw.draw_rectangle(
                slide=master_page, x=5, y=7, width=round(sz.width / 6), height=round(sz.height / 6)
            )
            _ = Draw.draw_text(
                slide=master_page, msg="Default Master Page", x=10, y=15, width=100, height=10, font_size=24
            )

            # set slide 1 to use the master page's slide number
            # but its own footer text
            slide1 = Draw.get_slide(doc=doc, idx=0)
            Draw.title_slide(slide=slide1, title="Slide 1")

            # IsPageNumberVisible = True: use the master page's slide number
            # change the master page's footer for first slide; does not work if the master already has a footer
            Props.set(slide1, IsPageNumberVisible=True, IsFooterVisible=True, FooterText="MU Slides")

            # add three more slides, which use the master page's
            # slide number and footer
            for i in range(1, 4):  # 1, 2, 3
                slide = Draw.insert_slide(doc=doc, idx=i)
                _ = Draw.bullets_slide(slide=slide, title=f"Slide {i}")
                Props.set(slide, IsPageNumberVisible=True, IsFooterVisible=True)

            # create master page 2
            master2 = Draw.insert_master_page(doc=doc, idx=1)
            _ = Draw.add_slide_number(master2)

            print("Master Page 2")
            Draw.show_shapes_info(master2)

            # link master page 2 to third slide
            Draw.set_master_page(slide=Draw.get_slide(doc=doc, idx=2), page=master2)

            # put ellipse and text on master page 2
            ellipse = Draw.draw_ellipse(
                slide=master2, x=5, y=7, width=round(sz.width / 6), height=round(sz.height / 6)
            )
            Props.set(ellipse, FillColor=CommonColor.GREEN_YELLOW)
            _ = Draw.draw_text(slide=master2, msg="Master Page 2", x=10, y=15, width=100, height=10, font_size=24)

            GUI.set_visible(is_visible=True, odoc=doc)

            Lo.delay(2_000)

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
