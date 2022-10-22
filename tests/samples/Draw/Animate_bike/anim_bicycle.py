from __future__ import annotations

import uno
from com.sun.star.drawing import XDrawPage

from ooodev.utils.lo import Lo
from ooodev.office.draw import Draw, PolySides
from ooodev.utils.info import Info
from ooodev.utils.props import Props
from ooodev.utils.gui import GUI
from ooodev.utils.color import CommonColor
from ooodev.utils.type_var import PathOrStr
from ooo.dyn.drawing.circle_kind import CircleKind


class AnimBicycle:
    def __init__(self, fnm_bike: PathOrStr) -> None:
        self._fnm_bike = fnm_bike

    def animate(self) -> None:
        with Lo.Loader(Lo.ConnectPipe()) as loader:
            doc = Draw.create_draw_doc(loader)

            slide = Draw.get_slide(doc=doc, idx=0)
            if Draw.is_impress(doc):
                Draw.title_only_slide(slide=slide, header="Bicycle in Motion")

            GUI.set_visible(is_visible=True, odoc=doc)

            slide_size = Draw.get_slide_size(slide)

            square = Draw.draw_polygon(slide=slide, x=125, y=125, sides=PolySides(4), radius=25)
            Props.set(square, FillColor=CommonColor.LIGHT_GREEN)

            # default radius of 20, no. of sides
            pentagon = Draw.draw_polygon(slide=slide, x=150, y=75, sides=PolySides(5))
            Props.set(pentagon, FillColor=CommonColor.PURPLE)

            xs = (10, 30, 10, 30)
            ys = (10, 100, 100, 10)

            Draw.draw_lines(slide=slide, xs=xs, ys=ys)

            pie = Draw.draw_ellipse(slide=slide, x=30, y=slide_size.Width - 100, width=40, height=20)
            Props.set(
                pie,
                FillColor=CommonColor.LIGHT_SKY_BLUE,
                CircleStartAngle=9_000,  #   90 degrees ccw
                CircleEndAngle=36_000,  #    360 degrees ccw
                CircleKind=CircleKind.SECTION,
            )

            self._animate_bike(slide=slide)
            Lo.wait_enter()
            Lo.close_doc(doc)

    def _animate_bike(self, slide: XDrawPage) -> None:
        # fnm = Info.get_gallery_dir() / "transportation" / "Bicycle-Blue.png"

        shape = Draw.draw_image(slide=slide, fnm=self._fnm_bike, x=60, y=100, width=90, height=50)

        pt = Draw.get_position(shape)
        angle = Draw.get_rotation(shape)
        print(f"Start Angle: {int(angle)}")
        for i in range(19):
            Draw.set_position(shape=shape, x=pt.X + (i * 5), y=pt.Y)  # move right
            Draw.set_rotation(shape=shape, angle=angle + (i * 5))  # rotates ccw
            Lo.delay(200)

        print(f"Final Angle: {int(Draw.get_rotation(shape))}")
        Draw.print_matrix(Draw.get_transformation(shape))
