from __future__ import annotations

from ooodev.office.draw import Draw
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr


class BasicShow:
    def __init__(self, fnm: PathOrStr) -> None:
        if not FileIO.is_valid_path_or_str(fnm):
            raise ValueError(f'fnm is not a valid format for PathOrStr: "{fnm}"')
        p_fnm = FileIO.get_absolute_path(fnm)
        if not p_fnm.exists():
            raise FileNotFoundError(f"File fnm does not exist: {p_fnm}")
        if not p_fnm.is_file():
            raise ValueError(f'fnm is not a file: "{p_fnm}"')
        self._fnm = p_fnm

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
