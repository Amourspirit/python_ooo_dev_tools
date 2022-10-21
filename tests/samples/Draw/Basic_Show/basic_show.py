from ooodev.office.draw import Draw
from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.utils.props import Props
from ooodev.utils import type_var


class BasicShow:
    def __init__(self, fnm: type_var.PathOrStr) -> None:
        self._fnm = fnm
        self._is_automatic = False
        self._is_full_screen = True
        self._is_transiton_on_click = True
        self._start_with_navigator = False
        self._is_always_on_top = False
        self._is_endless = False
        self._pause = 0

    def show(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())
        try:
            doc = Lo.open_doc(fnm=self._fnm, loader=loader)
            try:
                # slideshow start() crashes if the doc is not visible
                GUI.set_visible(is_visible=True, odoc=doc)

                show = Draw.get_show(doc=doc)
                self._set_show_prop(show)
                Props.show_obj_props("Slide show", show)

                show.start()

                sc = Draw.get_show_controller(show)
                Draw.wait_ended(sc)

            finally:
                Lo.close_doc(doc)
        finally:
            Lo.close_office()

    # region Properties
    @property
    def is_automatic(self) -> bool:
        """Specifies is_automatic"""
        return self._is_automatic

    @is_automatic.setter
    def is_automatic(self, value: bool):
        self._is_automatic = value

    @property
    def is_full_screen(self) -> bool:
        """Specifies is_full_screen"""
        return self._is_full_screen

    @is_full_screen.setter
    def is_full_screen(self, value: bool):
        self._is_full_screen = value

    @property
    def is_transiton_on_click(self) -> bool:
        """Specifies is_transiton_on_click"""
        return self._is_transiton_on_click

    @is_transiton_on_click.setter
    def is_transiton_on_click(self, value: bool):
        self._is_transiton_on_click = value

    @property
    def start_with_navigator(self) -> str:
        """Specifies start_with_navigator"""
        return self._start_with_navigator

    @start_with_navigator.setter
    def start_with_navigator(self, value: str):
        self._start_with_navigator = value

    @property
    def is_always_on_top(self) -> bool:
        """Specifies is_always_on_top"""
        return self._is_always_on_top

    @is_always_on_top.setter
    def is_always_on_top(self, value: bool):
        self._is_always_on_top = value

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

    # endregion Properties

    # region Set Slideshow properties
    def _set_show_prop(self, show: object) -> None:
        Props.set(
            show,
            IsAutomatic=self._is_automatic,
            IsFullScreen=self._is_full_screen,
            IsTransitionOnClick=self._is_transiton_on_click,
            StartWithNavigator=self._start_with_navigator,
            IsAlwaysOnTop=self._is_always_on_top,
            IsEndless=self._is_endless,
            Pause=self._pause,
        )

    # endregion Set Slideshow properties
