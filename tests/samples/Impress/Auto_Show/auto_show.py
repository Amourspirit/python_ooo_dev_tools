from __future__ import annotations

import uno
from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.draw import Draw, DrawingSlideShowKind
from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.presentation.animation_speed import AnimationSpeed as AnimationSpeed
from ooo.dyn.presentation.fade_effect import FadeEffect as FadeEffect


class AutoShow:
    def __init__(self, fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=fnm, raise_err=True)
        self._fnm = fnm
        self._is_endless = False
        self._pause = 0
        self._duration = 1
        self._end_delay = 2000
        self._fade_effect = FadeEffect.NONE

    def show(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = Lo.open_doc(self._fnm, loader)

            # slideshow start() crashes if the doc is not visible
            GUI.set_visible(is_visible=True, odoc=doc)

            # set up a fast automatic change between all the slides
            slides = Draw.get_slides_list(doc)
            for slide in slides:
                Draw.set_transition(
                    slide=slide,
                    fade_effect=self._fade_effect,
                    speed=AnimationSpeed.FAST,
                    change=DrawingSlideShowKind.AUTO_CHANGE,
                    duration=self._duration,
                )

            show = Draw.get_show(doc)
            Props.show_obj_props("Slide Show", show)
            self._set_show_prop(show)
            # Props.set(show, IsEndless=True, Pause=0)

            Lo.delay(500)
            Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)
            # show.start() starts slideshow but not necessarily in 100% full screen
            # show.start()

            sc = Draw.get_show_controller(show)
            Draw.wait_last(sc=sc, delay=self._end_delay)
            Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION_END)
            Lo.delay(500)

            msg_result = MsgBox.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                print("Ending the slide show")
                Lo.close_doc(doc=doc, deliver_ownership=True)
                Lo.close_office()
            else:
                print("Keeping document open")
        except Exception:
            Lo.close_office()
            raise

    def _set_show_prop(self, show: object) -> None:
        Props.set(show, IsEndless=self._is_endless, Pause=self._pause)

    @property
    def is_endless(self) -> bool:
        """Specifies is_endless"""
        return self._is_endless

    @is_endless.setter
    def is_endless(self, value: bool):
        self._is_endless = value

    @property
    def pause(self) -> int:
        """Specifies pause"""
        return self._pause

    @pause.setter
    def pause(self, value: int):
        self._pause = value

    @property
    def duration(self) -> int:
        """Specifies duration in seconds of each slide"""
        return self._duration

    @duration.setter
    def duration(self, value: int):
        self._duration = value

    @property
    def fade_effect(self) -> FadeEffect:
        """Specifies fade_effect"""
        return self._fade_effect

    @fade_effect.setter
    def fade_effect(self, value: FadeEffect):
        self._fade_effect = value

    @property
    def end_delay(self) -> int:
        """Specifies delay in seconds to wait  slideshow ends"""
        return round(self._end_delay / 1000)

    @end_delay.setter
    def end_delay(self, value: int):
        self._end_delay = value * 1000
