import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.gui.gui import GUI
from ooodev.utils.date_time_util import DateUtil


def test_make_table(loader, bond_movies_table: list):
    # test require Writer be visible
    visible = False
    delay = 0  # 1_000

    doc = Write.create_doc(loader)
    try:
        if visible:
            GUI.set_visible(visible, doc)

        cursor = Write.get_cursor(doc)

        Write.append_para(cursor, "Table of Bond Movies")
        Write.style_prev_paragraph(cursor, "Heading 1")
        Write.append_para(cursor, 'The following table comes form "bondMovies.txt"\n')
        # Lock display updating
        with Lo.ControllerLock():
            Write.add_table(cursor=cursor, table_data=bond_movies_table)
            Write.end_paragraph(cursor)

        Lo.delay(delay)
        Write.append(cursor, f"Timestamp: {DateUtil.time_stamp()}")
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc, False)
