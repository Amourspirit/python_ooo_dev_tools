import pytest
import sys

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows")
def test_num_style(loader):
    """
    This test requires Write to be visible.
    If not visible then Write.is_anything_selected() will return false every time.
    """
    from ooodev.loader.lo import Lo
    from ooodev.office.write import Write
    from ooodev.gui.gui import GUI

    # Not sure why but this stopped working on windows.
    # It works in windows if just this test is run, but, not in a group of tests.

    visible = True
    delay = 300
    doc = Write.create_doc(loader)
    try:
        if visible:
            GUI.set_visible(visible, doc)

        assert Write.is_anything_selected(doc) is False
        # must be a view cursor and not a text cursor.
        # text cursor do no make selection at a document level.
        cursor = Write.get_view_cursor(doc)

        Write.append_para(cursor, "The following points are important:")
        pos = Write.append_para(cursor, "Have a good breakfast")
        cursor.goRight(0, False)
        cursor.goLeft(21, True)
        assert Write.is_anything_selected(doc)
        cursor.goRight(21, False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc, False)
