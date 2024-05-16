from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_python_script(loader) -> None:
    from ooodev.calc import CalcDoc
    from ooodev.utils.string.str_list import StrList

    doc = CalcDoc.create_doc(loader)
    try:
        psa = doc.python_script
        assert psa is not None
        code = StrList(sep="\n")
        code.append("from __future__ import annotations")
        code.append()
        code.append("def say_hello() -> None:")
        with code.indented():
            code.append('print("Hello World!")')
        code.append()
        code_str = str(code)
        assert psa.is_valid_python(code_str)
        psa.write_file("MyFile", code_str, allow_override=True)
        psa.write_file("MyFile", code_str, allow_override=True)
        psa_code = psa.read_file("MyFile")
        assert psa_code == code_str

        assert psa.file_exist("MyFile")
        assert psa.delete_file("MyFile")
        assert not psa.file_exist("MyFile")

        code.append("from __future__ import annotations")
        code.append()
        code.append("def say_hello() -> None")  # missing colon
        with code.indented():
            code.append('print("Hello World!")')
        code.append()
        code_str = str(code)
        assert psa.is_valid_python(code_str) is False

    finally:
        doc.close_doc()
