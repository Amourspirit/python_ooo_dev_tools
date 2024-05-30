from __future__ import annotations
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.io.sfa.sfa import Sfa
from ooodev.utils.string.str_list import StrList


def test_sfa(loader, tmp_path) -> None:
    from ooodev.calc import CalcDoc

    doc = None
    try:
        pth = Path(tmp_path, "tes_sfa.ods")
        doc = CalcDoc.create_doc(loader)

        sfa = Sfa()
        root = f"vnd.sun.star.tdoc:/{doc.runtime_uid}/"
        new_dir = root + "new_dir"
        assert sfa.exists(new_dir) is False
        sfa.inst.create_folder(new_dir)
        assert sfa.exists(new_dir)
        assert sfa.inst.is_folder(new_dir)
        s_file = new_dir + "/new_file.txt"
        sfa.write_text_file(s_file, "Hello, World!")

        # check that file can be accessed before saving
        txt = sfa.read_text_file(s_file)
        assert txt == "Hello, World!"

        doc.save_doc(pth)
    finally:
        if doc is not None:
            doc.close()
            doc = None

    doc = None
    try:
        doc = CalcDoc.open_doc(pth)
        sfa = Sfa()
        root = f"vnd.sun.star.tdoc:/{doc.runtime_uid}/"
        new_dir = root + "new_dir"
        assert sfa.exists(new_dir)
        assert sfa.inst.is_folder(new_dir)
        s_file = new_dir + "/new_file.txt"
        txt = sfa.read_text_file(s_file)
        assert txt == "Hello, World!"
    finally:
        if doc is not None:
            doc.close()
            doc = None


def test_sfa_copy_file_to_doc(loader, tmp_path) -> None:
    from ooodev.calc import CalcDoc

    doc = None
    pth = Path(tmp_path, "tes_sfa_copy.ods")
    try:
        tmp_txt_file = Path(tmp_path, "test_data_text.txt")
        s = "Hello, World!"
        with open(tmp_txt_file, "w") as f:
            f.write(s)
        assert tmp_txt_file.exists()

        sfa = Sfa()
        doc = CalcDoc.create_doc(loader)

        root = f"vnd.sun.star.tdoc:/{doc.runtime_uid}/"
        new_dir = root + "new_dir"
        assert sfa.exists(new_dir) is False
        sfa.inst.create_folder(new_dir)
        assert sfa.exists(new_dir)
        assert sfa.inst.is_folder(new_dir)
        s_file = new_dir + "/new_file.txt"
        sfa.inst.copy(source_url=tmp_txt_file.as_uri(), dest_url=s_file)

        # check that file can be accessed before saving
        txt = sfa.read_text_file(s_file)
        assert txt == "Hello, World!"

        doc.save_doc(pth)
    finally:
        if doc is not None:
            doc.close()
            doc = None

    doc = None
    try:
        doc = CalcDoc.open_doc(pth)
        sfa = Sfa()
        root = f"vnd.sun.star.tdoc:/{doc.runtime_uid}/"
        new_dir = root + "new_dir"
        assert sfa.exists(new_dir)
        assert sfa.inst.is_folder(new_dir)
        s_file = new_dir + "/new_file.txt"
        txt = sfa.read_text_file(s_file)
        assert txt.strip() == "Hello, World!"
    finally:
        if doc is not None:
            doc.close()
            doc = None


def test_sfa_py(loader, tmp_path) -> None:
    from ooodev.calc import CalcDoc

    pth = Path(tmp_path, "tes_sfa_py.ods")
    doc = None
    try:
        code = StrList(sep="\n")
        code.append("from __future__ import annotations")
        code.append("import sys")
        code.append()
        code.append("def say_hello(*args) -> None:")
        with code.indented():
            code.append('print("Hello World!")')
        code.append()
        code_str = str(code)
        doc = CalcDoc.create_doc(loader)

        sfa = Sfa()
        root = f"vnd.sun.star.tdoc:/{doc.runtime_uid}/Scripts/python/"
        new_dir = root + "new_dir"
        assert sfa.exists(new_dir) is False
        sfa.inst.create_folder(new_dir)
        assert sfa.exists(new_dir)
        assert sfa.inst.is_folder(new_dir)
        s_file = new_dir + "/new_file.txt"
        sfa.write_text_file(s_file, code_str)
        doc.save_doc(pth)
    finally:
        if doc is not None:
            doc.close()
            doc = None
