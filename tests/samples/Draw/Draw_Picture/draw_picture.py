from __future__ import annotations

import uno
from com.sun.star.drawing import XDrawPage

from ooodev.exceptions.ex import ShapeError
from ooodev.office.draw import Draw, Intensity
from ooodev.utils.color import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props


class DrawPicture:
    def show(self) -> None:
        with Lo.Loader(Lo.ConnectPipe()) as loader:
            doc = Draw.create_draw_doc(loader)
            GUI.set_visible(is_visible=True, odoc=doc)
            Lo.delay(1_000)  # need delay or zoom may not occur
            GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

            curr_slide = Draw.get_slide(doc=doc, idx=0)
            self._draw_shapes(curr_slide=curr_slide)

            s = Draw.draw_formula(slide=curr_slide, formula="func e^{i %pi} + 1 = 0", x=70, y=20, width=75, height=40)
            # Draw.report_pos_size(s)

            self._anim_shapes(curr_slide=curr_slide)

            s = Draw.find_shape_by_name(curr_slide, "text1")
            Draw.report_pos_size(s)

            Lo.wait_enter()
            Lo.close_doc(doc=doc, deliver_ownership=True)

    def _draw_shapes(self, curr_slide: XDrawPage) -> None:
        line1 = Draw.draw_line(slide=curr_slide, x1=50, y1=50, x2=200, y2=200)
        Props.set(line1, LineColor=CommonColor.BLACK)
        Draw.set_dashed_line(shape=line1, is_dashed=True)

        # red ellipse; uses (x, y) width, height
        circle1 = Draw.draw_ellipse(slide=curr_slide, x=100, y=100, width=75, height=25)
        Props.set(circle1, FillColor=CommonColor.RED)

        # rectangle with different fills; uses (x, y) width, height
        rect1 = Draw.draw_rectangle(slide=curr_slide, x=70, y=100, width=75, height=25)
        Props.set(rect1, FillColor=CommonColor.LIME)

        text1 = Draw.draw_text(
            slide=curr_slide, msg="Hello LibreOffice", x=120, y=120, width=60, height=30, font_size=24
        )
        Props.set(text1, Name="text1")
        # Props.show_props("TextShape's Text Properties", Draw.get_text_properties(text1))

        # gray transparent circle; uses (x,y), radius
        circle2 = Draw.draw_circle(slide=curr_slide, x=40, y=150, radius=20)
        Props.set(circle2, FillColor=CommonColor.GRAY)
        Draw.set_transparency(shape=circle2, level=Intensity(25))

        # thick line; uses (x,y), angle clockwise from x-axis, length
        line2 = Draw.draw_polar_line(slide=curr_slide, x=60, y=200, degrees=45, distance=100)
        Props.set(line2, LineWidth=300)

    def _anim_shapes(self, curr_slide: XDrawPage) -> None:
        # two animations of a circle and a line
        # he animation loop is:
        #    redraw shape, delay, update shape position/size

        # reduce circle size and move to the right
        xc = 40
        yc = 150
        radius = 40
        circle = None
        for _ in range(20):
            # move right
            if circle is not None:
                curr_slide.remove(circle)
            circle = Draw.draw_circle(slide=curr_slide, x=xc, y=yc, radius=radius)

            Lo.delay(200)
            xc += 5
            radius *= 0.95

        x2 = 140
        y2 = 110
        line = None
        for _ in range(25):
            if line is not None:
                curr_slide.remove(line)
            line = Draw.draw_line(slide=curr_slide, x1=40, y1=100, x2=x2, y2=y2)
            x2 -= 4
            y2 -= 4
