from __future__ import annotations
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.write import WriteDoc
from ooodev.utils.helper.dot_dict import DotDict


def test_writer_custom_props(loader, tmp_path) -> None:
    # get_sheet is overload method.
    # testing each overload.
    doc = None
    try:
        pth = Path(tmp_path, "test_writer_custom_props.odt")
        doc = WriteDoc.create_doc(loader)

        dd = DotDict()
        dd.doc_name = "Who knows"
        dd.doc_num = 1
        dd.boolean = True
        dd.empty_str = ""
        dd.nothing = None
        doc.set_custom_properties(dd)

        dd = doc.get_custom_properties()
        assert dd.doc_name == "Who knows"
        assert dd.doc_num == 1
        assert dd.boolean is True
        assert dd.empty_str == ""
        assert dd.nothing is None

        doc.save_doc(pth)
        doc.close()

        doc = WriteDoc.open_doc(pth)
        dd = doc.get_custom_properties()
        assert dd.doc_name == "Who knows"
        assert dd.doc_num == 1
        assert dd.boolean is True
        assert dd.empty_str == ""
        assert dd.nothing is None

    finally:
        if doc:
            doc.close()
