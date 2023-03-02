# Generate a Hilbert curve of the specified level.
# Created using a series of rounded blue lines.
# Save as hilbert.png
# Position/size the window, resize the page view
# Usage:
#     run DrawHilbert 4
# Using '6' takes about 2 minutes to fully draw. It's
# fun to try once.
# Using '7' causes the code to mis-calculate incr, so the
# line drawing goes off the left side of the canvas. And it
# takes forever to do it!
from __future__ import annotations
import math

import uno
from com.sun.star.drawing import XDrawPage

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.utils.props import Props
from ooodev.utils.lo import Lo
from ooodev.office.draw import Draw
from ooodev.utils.color import CommonColor
from ooodev.utils.gui import GUI
from ooodev.proto.size_obj import SizeObj

from ooo.dyn.drawing.line_cap import LineCap


class DrawHilbert:
    def __init__(self, level: int) -> None:
        self._x = -1
        self._y = -1
        self._incr = -1
        self._delay = -1
        self._level = -1
        self._slide: XDrawPage = None
        self._level_str = str(level)

    def draw(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())
        try:
            doc = Draw.create_draw_doc(loader)
            rect = GUI.get_screen_size()
            GUI.set_pos_size(doc=doc, x=0, y=0, width=round(rect.Width / 2), height=round(rect.Height - 40))
            GUI.set_visible(is_visible=True, odoc=doc)
            Lo.delay(2000)  # need delay or zoom may not occur

            GUI.zoom_value(value=75)

            self._slide = Draw.get_slide(doc=doc, idx=0)
            self._start_hilbert(self._level_str, Draw.get_slide_size(self._slide))

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

    def _start_hilbert(self, level_str: str, slide_size: SizeObj) -> None:
        self._level = 1
        try:
            self._level = int(level_str)
        except ValueError:
            print("Level is not an integer; using 1")

        if self._level < 1:
            print("Level must be >= 1; using 1")

        # store smallest mm dimension
        sq_width = min(slide_size.Height, slide_size.Width)
        self._x = sq_width - 10
        self._y = 10
        self._incr = round((sq_width - 20) / ((math.pow(2, self._level) - 1)))
        if self._incr == 0:
            raise RuntimeError("Hilbert increment is 0")
        else:
            print(f"Line increment: {self._incr}")

        self._delay = round(1000 / (self._level * self._level))
        print(f"Delay: {self._delay}")

        self._a(self._level)

    def _a(self, i: int) -> None:
        if i > 0:
            self._d(i - 1)
            self._move_by(-self._incr, 0)
            self._a(i - 1)
            self._move_by(0, self._incr)
            self._a(i - 1)
            self._move_by(self._incr, 0)
            self._b(i - 1)
            Lo.delay(self._delay)

    def _b(self, i: int) -> None:
        if i > 0:
            self._c(i - 1)
            self._move_by(0, -self._incr)
            self._b(i - 1)
            self._move_by(self._incr, 0)
            self._b(i - 1)
            self._move_by(0, self._incr)
            self._a(i - 1)
            Lo.delay(self._delay)

    def _c(self, i: int) -> None:
        if i > 0:
            self._b(i - 1)
            self._move_by(self._incr, 0)
            self._c(i - 1)
            self._move_by(0, -self._incr)
            self._c(i - 1)
            self._move_by(-self._incr, 0)
            self._d(i - 1)
            Lo.delay(self._delay)

    def _d(self, i: int) -> None:
        if i > 0:
            self._a(i - 1)
            self._move_by(0, self._incr)
            self._d(i - 1)
            self._move_by(-self._incr, 0)
            self._d(i - 1)
            self._move_by(0, -self._incr)
            self._c(i - 1)
            Lo.delay(self._delay)

    def _move_by(self, x_step: int, y_step: int) -> None:
        x_old = self._x
        y_old = self._y
        self._x += x_step
        self._y += y_step

        line = Draw.draw_line(slide=self._slide, x1=x_old, y1=y_old, x2=self._x, y2=self._y)
        # LineWidth in 1/100 mm units
        # LineCap: round off the line
        Props.set(line, LineColor=CommonColor.BLUE, LineWidth=110, LineCap=LineCap.ROUND)
