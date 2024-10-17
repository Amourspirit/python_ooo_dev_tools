from __future__ import annotations
from typing import TYPE_CHECKING, cast
from enum import Enum


from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.draw import Draw, Angle, DrawingGradientKind, DrawingHatchingKind, DrawingBitmapKind
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.awt.gradient import Gradient as Gradient

if TYPE_CHECKING:
    # the following is only needed for typings.
    # from __future__ import annotations takes care of the rest
    from com.sun.star.drawing import FillProperties  # service
    from com.sun.star.drawing import XDrawPage


class GradientKind(str, Enum):
    FILL = "fill"
    GRADIENT = "gradient"
    GRADIENT_NAME = "name"
    GRADIENT_NAME_PROPS = "name_props"
    HATCHING = "hatch"
    BITMAP = "bitmap"
    BITMAP_FILE = "file"


class DrawGradient:
    def __init__(self, gradient_kind: GradientKind, gradient_fnm: PathOrStr = "") -> None:
        self._gradien_kind = gradient_kind

        self._gradient_fnm = gradient_fnm
        if self._gradien_kind == GradientKind.BITMAP_FILE:
            # file has to be valid when bitmap file
            _ = FileIO.is_exist_file(self._gradient_fnm, True)
        self._x = 93
        self._y = 100
        self._width = 30
        self._height = 60
        self._start_color = CommonColor.LIME
        self._end_color = CommonColor.RED
        self._angle = 0
        self._name_gradient = DrawingGradientKind.NEON_LIGHT
        self._hatch_gradient = DrawingHatchingKind.GREEN_30_DEGREES
        self._bitmap_gradient = DrawingBitmapKind.FLORAL

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = Draw.create_draw_doc(loader)
            GUI.set_visible(is_visible=True, odoc=doc)
            Lo.delay(1_000)  # need delay or zoom may not occur
            GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

            curr_slide = Draw.get_slide(doc=doc, idx=0)

            if self._gradien_kind == GradientKind.FILL:
                self._gradient_fill(curr_slide)
            elif self._gradien_kind == GradientKind.GRADIENT:
                self._gradient(curr_slide)
            elif self._gradien_kind == GradientKind.GRADIENT_NAME:
                self._gradient_name(curr_slide, False)
            elif self._gradien_kind == GradientKind.GRADIENT_NAME_PROPS:
                self._gradient_name(curr_slide, True)
            elif self._gradien_kind == GradientKind.HATCHING:
                self._gradient_hatching(curr_slide)
            elif self._gradien_kind == GradientKind.BITMAP:
                self._gradient_bitmap(curr_slide)
            elif self._gradien_kind == GradientKind.BITMAP_FILE:
                self._gradient_bitmap_file(curr_slide)

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

    def _gradient_fill(self, curr_slide: XDrawPage) -> None:

        # rectangle shape is also com.sun.star.drawing.FillProperties service
        # casting is only at design time and is not really necessary;
        # however it gives easy access with typing support for other properties
        rect1 = cast(
            "FillProperties",
            Draw.draw_rectangle(slide=curr_slide, x=self._x, y=self._y, width=self._width, height=self._height),
        )
        Props.set(rect1, FillColor=self._start_color)
        # other properties can be set
        # rect1.FillTransparence = 55

    def _gradient(self, curr_slide: XDrawPage) -> None:

        # rectangle shape is also com.sun.star.drawing.FillProperties service
        # casting is only at design time and is not really necessary;
        # however it gives easy access with typing support for other properties
        rect1 = cast(
            "FillProperties",
            Draw.draw_rectangle(slide=curr_slide, x=self._x, y=self._y, width=self._width, height=self._height),
        )
        Draw.set_gradient_color(
            shape=rect1, start_color=self._start_color, end_color=self._end_color, angle=Angle(self._angle)
        )
        # rect1.FillTransparence = 40

    def _gradient_name(self, curr_slide: XDrawPage, set_props: bool) -> None:

        # rectangle shape is also com.sun.star.drawing.FillProperties service
        # casting is only at design time and is not really necessary;
        # however it gives easy access with typing support for other properties
        rect1 = cast(
            "FillProperties",
            Draw.draw_rectangle(slide=curr_slide, x=self._x, y=self._y, width=self._width, height=self._height),
        )
        grad = Draw.set_gradient_color(shape=rect1, name=self._name_gradient)
        if set_props:
            # grad = cast("Gradient", Props.get(rect1, "FillGradient"))
            # print(grad)
            grad.Angle = self._angle * 10  # in 1/10 degree units
            grad.StartColor = self._start_color
            grad.EndColor = self._end_color
            Draw.set_gradient_properties(shape=rect1, grad=grad)
        # rect1.FillTransparence = 40

    def _gradient_hatching(self, curr_slide: XDrawPage) -> None:
        # rectangle shape is also com.sun.star.drawing.FillProperties service
        # casting is only at design time and is not really necessary;
        # however it gives easy access with typing support for other properties
        rect1 = cast(
            "FillProperties",
            Draw.draw_rectangle(slide=curr_slide, x=self._x, y=self._y, width=self._width, height=self._height),
        )
        Draw.set_hatch_color(shape=rect1, name=self._hatch_gradient)
        # rect1.FillTransparence = 40

    def _gradient_bitmap(self, curr_slide: XDrawPage) -> None:
        # rectangle shape is also com.sun.star.drawing.FillProperties service
        # casting is only at design time and is not really necessary;
        # however it gives easy access with typing support for other properties
        rect1 = cast(
            "FillProperties",
            Draw.draw_rectangle(slide=curr_slide, x=self._x, y=self._y, width=self._width, height=self._height),
        )
        Draw.set_bitmap_color(shape=rect1, name=self._bitmap_gradient)
        # rect1.FillTransparence = 40

    def _gradient_bitmap_file(self, curr_slide: XDrawPage) -> None:
        # rectangle shape is also com.sun.star.drawing.FillProperties service
        # casting is only at design time and is not really necessary;
        # however it gives easy access with typing support for other properties
        rect1 = Draw.draw_rectangle(slide=curr_slide, x=self._x, y=self._y, width=self._width, height=self._height)
        Draw.set_bitmap_file_color(shape=rect1, fnm=self._gradient_fnm)
        # rect1.FillTransparence = 40

    # region properties
    @property
    def angle(self) -> int:
        """Specifies angle"""
        return self._angle

    @angle.setter
    def angle(self, value: int):
        self._angle = value

    @property
    def x(self) -> int:
        """Specifies x"""
        return self._x

    @x.setter
    def x(self, value: int):
        self._x = value

    @property
    def y(self) -> int:
        """Specifies x"""
        return self._y

    @y.setter
    def y(self, value: int):
        self._y = value

    @property
    def width(self) -> int:
        """Specifies width"""
        return self._width

    @width.setter
    def width(self, value: int):
        self._width = value

    @property
    def height(self) -> int:
        """Specifies height"""
        return self._height

    @height.setter
    def height(self, value: int):
        self._height = value

    @property
    def start_color(self) -> int:
        """Specifies start_color"""
        return self._start_color

    @start_color.setter
    def start_color(self, value: int):
        self._start_color = value

    @property
    def end_color(self) -> int:
        """Specifies end_color"""
        return self._end_color

    @end_color.setter
    def end_color(self, value: int):
        self._end_color = value

    @property
    def name_gradient(self) -> DrawingGradientKind:
        """Specifies name_gradient"""
        return self._name_gradient

    @name_gradient.setter
    def name_gradient(self, value: DrawingGradientKind):
        self._name_gradient = value

    @property
    def hatch_gradient(self) -> DrawingHatchingKind:
        """Specifies hatch_gradient"""
        return self._hatch_gradient

    @hatch_gradient.setter
    def hatch_gradient(self, value: DrawingHatchingKind):
        self._hatch_gradient = value

    @property
    def bitmap_gradient(self) -> DrawingBitmapKind:
        """Specifies bitmap_gradient"""
        return self._bitmap_gradient

    @bitmap_gradient.setter
    def bitmap_gradient(self, value: DrawingBitmapKind):
        self._bitmap_gradient = value

    # endregion properties
