import os
import pytest
from pathlib import Path
from typing import cast

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.images_lo import ImagesLo


def test_extract_images(loader, copy_fix_writer, tmp_path_fn):
    test_doc = copy_fix_writer("build.odt")

    doc = Write.open_doc(test_doc, loader)
    try:
        pics = Write.get_text_graphics(doc)
        assert len(pics) == 2

        for i, pic in enumerate(pics):
            fnm = Path(tmp_path_fn, f"graphic{i}.png")
            ImagesLo.save_graphic(pic, fnm, "png")  # ".gif", "gif")
            assert fnm.exists()
    finally:
        Lo.close_doc(doc, False)
