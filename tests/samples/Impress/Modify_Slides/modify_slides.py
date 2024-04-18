from __future__ import annotations

import uno
from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.draw import Draw
from ooodev.utils.file_io import FileIO
from ooodev.gui.gui import GUI
from ooodev.utils.info import Info
from ooodev.loader.lo import Lo
from ooodev.utils.type_var import PathOrStr


class ModifySlides:
    def __init__(self, fnm: PathOrStr, im_fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=fnm, raise_err=True)
        _ = FileIO.is_exist_file(fnm=im_fnm, raise_err=True)
        self._fnm = FileIO.get_absolute_path(fnm)
        self._im_fnm = FileIO.get_absolute_path(im_fnm)

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = Lo.open_doc(self._fnm, loader)

            # slideshow start() crashes if the doc is not visible
            GUI.set_visible(is_visible=True, odoc=doc)

            if not Info.is_doc_type(obj=doc, doc_type=Lo.Service.IMPRESS):
                print("-- Not a slides presentation")
                Lo.close_office()
                return

            slides = Draw.get_slides(doc)
            num_slides = slides.getCount()
            print(f"No. of slides: {num_slides}")

            # add a title-only slide with a graphic at the end
            last_page = slides.insertNewByIndex(num_slides)
            Draw.title_only_slide(slide=last_page, header="Any Questions?")
            Draw.draw_image(slide=last_page, fnm=self._im_fnm)

            # add a title/subtitle slide at the start
            first_page = slides.insertNewByIndex(0)
            Draw.title_slide(slide=first_page, title="Interesting Slides", sub_title="Brought to you by ODEV")

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
