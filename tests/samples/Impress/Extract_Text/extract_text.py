# Attempts to extract the text from the slide deck.
# The order of the text extracted may not be the same as the order
# that it appears in the file; it depends on the order that the text shapes are saved inside the file.
from __future__ import annotations

from ooodev.utils.lo import Lo
from ooodev.office.draw import Draw
from ooodev.utils.type_var import PathOrStr
from ooodev.utils.file_io import FileIO


class ExtractText:
    def __init__(self, fnm: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=fnm, raise_err=True)
        self._fnm = FileIO.get_absolute_path(fnm)

    def extract(self) -> None:
        with Lo.Loader(Lo.ConnectPipe(headless=True)) as loader:
            doc = Lo.open_doc(fnm=self._fnm, loader=loader)

            if Draw.is_shapes_based(doc):
                print("Text Content".center(46, "-"))
                print(Draw.get_shapes_text(doc))
                print("-" * 46)
            else:
                print("Text extraction unsupported for this document type")

            Lo.delay(1000)
            Lo.close_doc(doc)
