from __future__ import annotations

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.draw import Draw, FadeEffect, AnimationSpeed, DrawingGradientKind, DrawingSlideShowKind
from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.props import Props

from ooo.dyn.presentation.animation_effect import AnimationEffect
from ooo.dyn.presentation.click_action import ClickAction


class SlideShow:
    def show(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        # create Impress page or Draw slide
        try:
            doc = Draw.create_impress_doc(loader)

            while Draw.get_slides_count(doc) < 3:
                _ = Draw.add_slide(doc)

            # ---- The first page
            slide = Draw.get_slide(doc=doc, idx=0)
            Draw.set_transition(
                slide=slide,
                fade_effect=FadeEffect.FADE_FROM_RIGHT,
                speed=AnimationSpeed.FAST,
                change=DrawingSlideShowKind.AUTO_CHANGE,
                duration=1,
            )
            # draw a square at the top left of the page; and text
            sq1 = Draw.draw_rectangle(slide=slide, x=10, y=10, width=50, height=50)
            Props.set(sq1, Effect=AnimationEffect.WAVYLINE_FROM_BOTTOM)
            # square appears in 'wave' (pixels by pixels)
            _ = Draw.draw_text(slide=slide, msg="Page 1", x=70, y=20, width=60, height=30, font_size=24)

            # ---- The second page
            slide = Draw.get_slide(doc=doc, idx=1)
            Draw.set_transition(
                slide=slide,
                fade_effect=FadeEffect.FADE_FROM_RIGHT,
                speed=AnimationSpeed.FAST,
                change=DrawingSlideShowKind.AUTO_CHANGE,
                duration=1,
            )
            # draw a circle at the bottom right of second page; and text
            circle1 = Draw.draw_ellipse(slide=slide, x=212, y=150, width=50, height=50)
            # hide circle after drawing
            Props.set(circle1, Effect=AnimationEffect.HIDE)

            _ = Draw.draw_text(slide=slide, msg="Page 2", x=170, y=170, width=60, height=30, font_size=24)

            name_slide = "page two"
            Draw.set_name(slide=slide, name=name_slide)

            # ---- The third page
            slide = Draw.get_slide(doc=doc, idx=2)
            Draw.set_transition(
                slide=slide,
                fade_effect=FadeEffect.ROLL_FROM_LEFT,
                speed=AnimationSpeed.MEDIUM,
                change=DrawingSlideShowKind.AUTO_CHANGE,
                duration=2,
            )
            _ = Draw.draw_text(slide=slide, msg="Page 3", x=120, y=75, width=60, height=30, font_size=24)
            # draw a circle containing text
            circle2 = Draw.draw_ellipse(slide=slide, x=10, y=8, width=50, height=50)
            Draw.add_text(shape=circle2, msg="Click to go\nto Page1")
            Draw.set_gradient_color(shape=circle2, name=DrawingGradientKind.MAHOGANY)

            # clicking makes the presentation jump to page one
            Props.set(circle2, Effect=AnimationEffect.FADE_FROM_BOTTOM, OnClick=ClickAction.FIRSTPAGE)

            # draw a square with text
            sq2 = Draw.draw_rectangle(slide=slide, x=220, y=8, width=50, height=50)
            Draw.add_text(shape=sq2, msg="Click to go\nto Page 2")
            Draw.set_gradient_color(shape=sq2, name=DrawingGradientKind.MAHOGANY)

            # clicking makes the presentation jump to page two by using a bookmark
            Props.set(sq2, Effect=AnimationEffect.FADE_FROM_BOTTOM, OnClick=ClickAction.BOOKMARK, Bookmark=name_slide)

            # slideshow start() crashes if the doc is not visible
            GUI.set_visible(is_visible=True, odoc=doc)
            show = Draw.get_show(doc)

            # a full-screen slide show
            Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)
            Lo.delay(500)
            # show.start() starts slideshow but not necessarily in 100% full screen
            # show.start()

            Props.show_obj_props("Slide show", show)
            sc = Draw.get_show_controller(show)

            # Draw.wait_last(sc, 2000)
            # Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION_END)
            Draw.wait_ended(sc)
            Lo.delay(1_000)
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
