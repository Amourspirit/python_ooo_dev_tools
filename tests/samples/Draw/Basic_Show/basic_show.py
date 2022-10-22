from ooodev.office.draw import Draw
from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.utils.props import Props
from ooodev.utils import type_var


class BasicShow:
    def __init__(self, fnm: type_var.PathOrStr) -> None:
        self._fnm = fnm

    def show(self) -> None:
        with Lo.Loader(Lo.ConnectPipe()) as loader:
            doc = Lo.open_doc(fnm=self._fnm, loader=loader)
            try:
                # slideshow start() crashes if the doc is not visible
                GUI.set_visible(is_visible=True, odoc=doc)

                show = Draw.get_show(doc=doc)
                Props.show_obj_props("Slide show", show)

                show.start()

                sc = Draw.get_show_controller(show)
                Draw.wait_ended(sc)

            finally:
                Lo.close_doc(doc)
