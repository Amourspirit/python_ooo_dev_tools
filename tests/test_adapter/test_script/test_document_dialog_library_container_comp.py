from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


from ooodev.adapter.script.document_dialog_library_container_comp import DocumentDialogLibraryContainerComp


def test_instance(loader) -> None:
    from ooodev.calc import CalcDoc

    inst = DocumentDialogLibraryContainerComp.from_lo()
    assert inst is not None
    doc = CalcDoc.create_doc()
    inst2 = doc.dialog_libraries
    assert inst2 is not None
    doc.close()
