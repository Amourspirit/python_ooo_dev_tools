import pytest
import sys

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.lo import Lo
from ooodev.write import Write
from ooodev.write import WriteDoc

# on windows getting Fatal Python error: Aborted even though the test runs fine when run by itself.


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows in a group")
def test_num_style(loader):
    """
    This test requires Write to be visible.
    If not visible then Write.is_anything_selected() will return false every time.
    """

    visible = True
    delay = 300
    doc = WriteDoc(Write.create_doc(loader))
    try:
        if visible:
            doc.set_visible(visible=visible)

        assert doc.is_anything_selected() is False
        # must be a view cursor and not a text cursor.
        # text cursor do no make selection at a document level.
        cursor = doc.get_view_cursor()

        cursor.append_para("The following points are important:")
        cursor.append_para("Have a good breakfast")
        cursor.go_right(0)
        cursor.go_left(21, True)
        assert doc.is_anything_selected()
        cursor.go_right(21)

        Lo.delay(delay)
    finally:
        doc.close_doc()
