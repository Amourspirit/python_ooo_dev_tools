from __future__ import annotations
from pathlib import Path

import uno
from com.sun.star.text import XText
from com.sun.star.lang import XComponent

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.draw import Draw
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.loader.lo import Lo
from ooodev.utils.type_var import PathOrStr


class PointsBuilder:
    def __init__(self, points_fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=points_fnm, raise_err=True)
        self._points_fnm = FileIO.get_absolute_path(points_fnm)

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        # create Impress page or Draw slide
        try:
            self._report_templates()
            tmpl_name = "Piano.otp"  # "Inspiration.otp"
            template_fnm = Path(Draw.get_slide_template_path(), tmpl_name)
            _ = FileIO.is_exist_file(template_fnm, True)
            doc = Lo.create_doc_from_template(template_path=template_fnm, loader=loader)

            self._read_points(doc)

            print(f"Total no. of slides: {Draw.get_slides_count(doc)}")

            GUI.set_visible(is_visible=True, odoc=doc)
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

    def _report_templates(self) -> None:
        template_dirs = Info.get_dirs(setting="Template")
        print("Templates dir:")
        for dir in template_dirs:
            print(f"  {dir}")

        temmplate_dir = Draw.get_slide_template_path()
        print()
        print(f'Templates files in "{temmplate_dir}"')
        template_fnms = FileIO.get_file_paths(temmplate_dir)
        for fnm in template_fnms:
            print(f"  {fnm}")

    def _read_points(self, doc: XComponent) -> None:
        # Read in a text file of points which are converted to slides.
        # Formatting rules:
        # * ">", ">>", etc are points and their levels
        # * any other lines are the title text of a new slide
        try:

            def process_bullet(line: str, xbody: XText) -> None:
                # count the number of '>'s to determine the bullet level
                if xbody is None:
                    print(f"No slide body for {line}")
                    return

                pos = 0
                s_lst = [*line]
                ch = s_lst[pos]
                while ch == ">":
                    pos += 1
                    ch = s_lst[pos]
                sub_str = "".join(s_lst[pos:]).strip()
                Draw.add_bullet(bulls_txt=xbody, level=pos - 1, text=sub_str)

            body: XText = None
            with open(self._points_fnm, "r") as file:
                # remove empty lines
                data = (row for row in file if row.strip())
                # chain generator
                # strip of remove anything starting //
                # // for comment
                data = (row for row in data if not row.lstrip().startswith("//"))

                for row in data:
                    ch = row[:1]
                    if ch == ">":
                        process_bullet(line=row, xbody=body)
                    else:
                        curr_slide = Draw.add_slide(doc)
                        body = Draw.bullets_slide(slide=curr_slide, title=row)
            print(f"Read in point file: {self._points_fnm.name}")
        except Exception as e:
            print(f"Error reading points file: {self._points_fnm}")
            print(f"  {e}")
