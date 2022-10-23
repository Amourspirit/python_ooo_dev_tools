from __future__ import annotations

import uno
from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.draw import Draw
from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.utils.props import Props
from ooodev.utils.file_io import FileIO
from ooodev.utils.type_var import PathOrStr


class CustomShow:
    def __init__(self, fnm: PathOrStr, *slide_idx: int) -> None:
        if not FileIO.is_valid_path_or_str(fnm):
            raise ValueError(f'fnm is not a valid format for PathOrStr: "{fnm}"')
        p_fnm = FileIO.get_absolute_path(fnm)
        if not p_fnm.exists():
            raise FileNotFoundError(f"File fnm does not exist: {p_fnm}")
        if not p_fnm.is_file():
            raise ValueError(f'fnm is not a file: "{p_fnm}"')
        self._fnm = p_fnm
        for idx in slide_idx:
            if idx < 0:
                raise IndexError("Index cannot be negative")
        self._idxs = slide_idx

    def show(self) -> None:
        try:
            loader = Lo.load_office(Lo.ConnectPipe())
        except Exception:
            Lo.close_office()
            raise
        doc = Lo.open_doc(fnm=self._fnm, loader=loader)
        try:
            # slideshow start() crashes if the doc is not visible
            GUI.set_visible(is_visible=True, odoc=doc)

            if len(self._idxs) > 0:
                _ = Draw.build_play_list(doc, "ShortPlay", *self._idxs)
                show = Draw.get_show(doc=doc)
                Props.set(show, CustomShow="ShortPlay")
                Props.show_obj_props("Slide show", show)
                show.start()
                sc = Draw.get_show_controller(show)
                Draw.wait_ended(sc)

                Lo.delay(2000)
                msg_result = MsgBox.msgbox(
                    "Do you wish to close document?",
                    "All done",
                    boxtype=MessageBoxType.QUERYBOX,
                    buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
                )
                if msg_result == MessageBoxResultsEnum.YES:
                    Lo.close_doc(doc=doc)
                else:
                    print("Keeping document open")
            else:
                MsgBox.msgbox(
                    "There were no slides indexes to create a slide show.",
                    "No Slide Indexes",
                    boxtype=MessageBoxType.WARNINGBOX,
                )

        except Exception:
            Lo.close_doc(doc=doc)
            raise
