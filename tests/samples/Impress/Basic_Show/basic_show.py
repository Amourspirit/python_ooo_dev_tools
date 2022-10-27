from __future__ import annotations

import uno
from ooodev.office.draw import Draw
from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr


class BasicShow:
    def __init__(self, fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=fnm, raise_err=True)
        self._fnm = FileIO.get_absolute_path(fnm)

    def show(self) -> None:
        with Lo.Loader(Lo.ConnectPipe()) as loader:
            doc = Lo.open_doc(fnm=self._fnm, loader=loader)
            try:
                # slideshow start() crashes if the doc is not visible
                GUI.set_visible(is_visible=True, odoc=doc)

                show = Draw.get_show(doc=doc)
                Props.show_obj_props("Slide show", show)

                Lo.delay(500)
                Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)
                # show.start() starts slideshow but not necessarily in 100% full screen
                # show.start()

                sc = Draw.get_show_controller(show)
                Draw.wait_ended(sc)

            finally:
                Lo.close_doc(doc)
