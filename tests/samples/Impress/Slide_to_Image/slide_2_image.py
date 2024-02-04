from __future__ import annotations
import tempfile
from pathlib import Path

import uno
from ooodev.office.draw import Draw
from ooodev.utils.file_io import FileIO
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.info import Info
from ooodev.loader.lo import Lo
from ooodev.utils.type_var import PathOrStr


class Slide2Image:
    def __init__(self, fnm: PathOrStr, idx: int, img_fmt: str, out_dir: PathOrStr = "") -> None:
        _ = FileIO.is_exist_file(fnm, True)
        self._fnm = FileIO.get_absolute_path(fnm)
        if out_dir:
            _ = FileIO.is_exist_dir(out_dir, True)
            self._out_dir = FileIO.get_absolute_path(out_dir)
        else:
            self._out_dir = Path(tempfile.mkdtemp())
        if idx < 0:
            Lo.print("Index is less then zero. Using zero")
            idx = 0
        self._idx = idx
        self._img_fmt = img_fmt.strip()

    def main(self) -> None:
        # connect headless. will not need to see slides
        with Lo.Loader(Lo.ConnectPipe(headless=True)) as loader:
            doc = Lo.open_doc(fnm=self._fnm, loader=loader)

            if not Info.is_doc_type(obj=doc, doc_type=Lo.Service.IMPRESS):
                Lo.print("-- Not a slides presentation")
                return

            slide = Draw.get_slide(doc=doc, idx=self._idx)

            names = ImagesLo.get_mime_types()
            Lo.print("Known GraphicExportFilter mime types:")
            for name in names:
                Lo.print(f"  {name}")

            out_fnm = self._out_dir / f"{self._fnm.stem}{self._idx}.{self._img_fmt}"
            Lo.print(f'Saving page {self._idx} to "{out_fnm}"')
            mime = ImagesLo.change_to_mime(self._img_fmt)
            Draw.save_page(page=slide, fnm=out_fnm, mime_type=mime)
            Lo.close_doc(doc)
