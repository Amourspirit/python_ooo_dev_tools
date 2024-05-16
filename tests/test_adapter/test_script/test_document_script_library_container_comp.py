from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


from ooodev.adapter.script.document_script_library_container_comp import DocumentScriptLibraryContainerComp


def test_instance(loader) -> None:
    from ooodev.write import WriteDoc

    inst = DocumentScriptLibraryContainerComp.from_lo()
    # from ooodev.calc import CalcDoc
    assert inst is not None

    # doc = CalcDoc.create_doc()
    doc = WriteDoc.create_doc()
    inst2 = doc.basic_libraries
    assert inst2 is not None
    doc.close()
