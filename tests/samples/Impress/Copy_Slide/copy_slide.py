from __future__ import annotations
from typing import TYPE_CHECKING


import uno
from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.draw import Draw
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr

if TYPE_CHECKING:
    # the following is only needed for typings.
    # from __future__ import annotations takes care of the rest
    from com.sun.star.lang import XComponent


class CopySlide:
    def __init__(self, fnm: PathOrStr, from_idx: int, to_idx: int) -> None:
        idx1 = int(from_idx)
        idx2 = int(to_idx)
        if idx1 < 0:
            raise IndexError("From Index must be greater or equal to 0")
        if idx2 < 0:
            raise IndexError("To Index must be greater or equal to 0")
        self._from_idx = idx1
        self._to_idx = idx2
        if not FileIO.is_valid_path_or_str(fnm):
            raise ValueError(f'fnm is not a valid format for PathOrStr: "{fnm}"')
        p_fnm = FileIO.get_absolute_path(fnm)
        if not p_fnm.exists():
            raise FileNotFoundError(f"File fnm does not exist: {p_fnm}")
        if not p_fnm.is_file():
            raise ValueError(f'fnm is not a file: "{p_fnm}"')
        self._fnm = p_fnm

    def copy(self) -> None:
        try:
            loader = Lo.load_office(Lo.ConnectPipe())
        except Exception:
            Lo.close_office()
            raise

        doc = Lo.open_doc(fnm=self._fnm, loader=loader)
        try:
            num_slides = Draw.get_slides_count(doc)
            if self._from_idx >= num_slides or self._to_idx >= num_slides:
                Lo.close_office()
                raise IndexError("One or both indicies are out of range")

            GUI.set_visible(is_visible=True, odoc=doc)

            self._copy_to(doc=doc)

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
        except IndexError:
            raise
        except Exception:
            Lo.close_doc(doc=doc)
            raise

    def _copy_to(self, doc: XComponent) -> None:
        # Copy fromIdx slide to the clipboard in slide-sorter mode,
        # then paste it to after the toIdx slide.

        ctrl = GUI.get_current_controller(doc)

        # Switch to slide sorter view so that slides can be pasted
        Lo.dispatch_cmd(cmd="DiaMode")

        # give Office 5 seconds of time to do it
        Lo.delay(5000)

        from_slide = Draw.get_slide(doc, self._from_idx)
        to_slide = Draw.get_slide(doc, self._to_idx)

        Draw.goto_page(ctrl, from_slide)
        Lo.dispatch_cmd(cmd="Copy")
        print(f"Copied {self._from_idx}")

        # elect this slide; 'doc' version of gotoPage() pastes incorrectly
        Draw.goto_page(ctrl, to_slide)
        Lo.dispatch_cmd("Paste")
        print(f"Paste to after {self._to_idx}")

        # Lo.dispatchCmd("PageMode");  // back to normal mode (not working)
        Lo.dispatch_cmd(cmd="DrawingMode")
