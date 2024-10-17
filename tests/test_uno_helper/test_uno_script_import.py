from __future__ import annotations
import sys
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.write import WriteDoc
from ooodev.calc import CalcDoc
from ooodev.uno_helper.importer import importer_shared_script
from ooodev.uno_helper.importer import importer_doc_script


# angle implements BaseIntValue so this test all dunder methods.

# https://github.com/LibreOffice/core/blob/8950ecd0718fdd5b773922b223d28deefe9adf59/comphelper/source/processfactory/processfactory.cxx


def test_import_shared(loader) -> None:
    # LibreOffice has a python script named Capitalise in the shared folder
    doc = None
    try:
        with importer_shared_script():
            import Capitalise  # noqa # type: ignore
        assert "Capitalise" in sys.modules
        del sys.modules["Capitalise"]
    finally:
        if doc:
            doc.close()


def test_import_user(loader) -> None:
    doc = WriteDoc.create_doc(loader=loader)
    pth = doc.python_script.get_user_script_path()
    assert pth


def test_import_doc(loader, copy_fix_calc) -> None:
    # the calc_embedded_mod_hello doc contains a module mod_hello /Scripts/python/mod_hello.py
    # using the importer_doc_script_context() context manager, we can import the module
    doc = None
    try:
        fnm = copy_fix_calc("calc_embedded_mod_hello.ods")
        doc = CalcDoc.open_doc(fnm=fnm, loader=loader)
        with importer_doc_script():
            import mod_hello  # type: ignore
        assert "mod_hello" in sys.modules
        assert hasattr(mod_hello, "say_hello")
        del sys.modules["mod_hello"]
    finally:
        if doc:
            doc.close()
