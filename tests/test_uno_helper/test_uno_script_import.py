from __future__ import annotations
import sys
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader import Lo
from ooodev.write import WriteDoc
from ooodev.uno_helper.importer import importer_shared_script_context


# angle implements BaseIntValue so this test all dunder methods.

# https://github.com/LibreOffice/core/blob/8950ecd0718fdd5b773922b223d28deefe9adf59/comphelper/source/processfactory/processfactory.cxx


def test_import_shared(loader) -> None:
    doc = None
    try:
        with importer_shared_script_context():
            import Capitalise
        assert "Capitalise" in sys.modules
        del sys.modules["Capitalise"]
    finally:
        if doc:
            doc.close()


def test_import_user(loader) -> None:
    doc = WriteDoc.create_doc(loader=loader)
    pth = doc.python_script.get_user_script_path()
    assert pth
