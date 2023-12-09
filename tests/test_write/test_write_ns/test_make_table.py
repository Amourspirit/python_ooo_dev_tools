import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.lo import Lo
from ooodev.write import Write
from ooodev.write import WriteDoc
from ooodev.utils.date_time_util import DateUtil


def test_make_table(loader, bond_movies_table: list):
    # test require Writer be visible
    visible = False
    delay = 0  # 1_000

    doc = WriteDoc(Write.create_doc(loader))
    try:
        if visible:
            doc.set_visible()

        cursor = doc.get_cursor()

        cursor.append_para("Table of Bond Movies")
        cursor.style_prev_paragraph("Heading 1")
        cursor.append_para('The following table comes form "bondMovies.txt"\n')
        # Lock display updating
        with Lo.ControllerLock():
            cursor.add_table(table_data=bond_movies_table)
            cursor.end_paragraph()

        Lo.delay(delay)
        cursor.append(f"Timestamp: {DateUtil.time_stamp()}")
        Lo.delay(delay)
    finally:
        doc.close_doc()
