from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_calc_script(loader) -> None:
    from ooodev.calc import CalcDoc

    doc = None
    try:
        doc = CalcDoc.create_doc()
        py_script = doc.python_script
        assert py_script
        scp = py_script.shared_script_provider
        assert scp is not None
        assert scp.uno_packages_sp
    finally:
        if doc:
            doc.close()
