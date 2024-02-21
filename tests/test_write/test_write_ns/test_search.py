import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.lo import Lo
from ooodev.write import WriteDoc


def test_search(loader):
    """
    This test requires Write to be visible.
    If not visible then Write.is_anything_selected() will return false every time.
    """

    visible = True
    delay = 300
    doc = WriteDoc.create_doc(loader)
    try:
        if visible:
            doc.set_visible(visible=visible)

        cursor = doc.get_cursor()

        cursor.append_para("The following points are important:")
        cursor.append_para("Have a good breakfast")

        search_desc = doc.create_search_descriptor()
        search_desc.set_search_string("important")
        search_desc.search_regular_expression = False
        first = doc.find_first(search_desc.component)
        assert first is not None

        Lo.delay(delay)
    finally:
        doc.close_doc()
