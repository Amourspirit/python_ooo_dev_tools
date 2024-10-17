import pytest
from pathlib import Path

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


from ooodev.write import Write
from ooodev.write import WriteDoc
from ooodev.utils.images_lo import ImagesLo


def test_extract_images(loader, copy_fix_writer, tmp_path_fn):
    test_doc = copy_fix_writer("build.odt")

    doc = WriteDoc(Write.open_doc(test_doc, loader))
    try:
        pics = doc.get_text_graphics()
        assert len(pics) == 2

        for i, pic in enumerate(pics):
            fnm = Path(tmp_path_fn, f"graphic{i}.png")
            ImagesLo.save_graphic(pic, fnm, "png")  # ".gif", "gif")
            assert fnm.exists()
    finally:
        doc.close_doc()
