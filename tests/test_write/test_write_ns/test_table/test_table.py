import pytest

if __name__ == "__main__":
    pytest.main([__file__])

# pylint: disable=import-outside-toplevel
# pylint: disable=wrong-import-position

from ooodev.write import WriteDoc


def test_make_table(loader, bond_movies_table: list):

    doc = WriteDoc.create_doc(loader)
    try:

        cursor = doc.get_cursor()

        cursor.append_para("Table of Bond Movies")
        cursor.append_para('The following table comes form "bondMovies.txt"\n')
        # Lock display updating
        with doc:
            cursor.add_table(table_data=bond_movies_table)
            cursor.end_paragraph()

        assert len(doc.tables) == 1

    finally:
        doc.close()


def test_make_single_cell_table(loader):

    doc = WriteDoc.create_doc(loader)
    try:
        cursor = doc.get_cursor()

        tbl = cursor.add_table(table_data=(("Hello",),))
        assert len(tbl.rows) == 1
        assert len(tbl.columns) == 1
        cell = tbl["A1"]
        assert cell.cell_name == "A1"
        assert cell.value == "Hello"
    finally:
        doc.close()


def test_iter_cells(loader, copy_fix_writer):
    fnm = copy_fix_writer("bondMovies.odt")
    doc = WriteDoc.open_doc(fnm=fnm, loader=loader)
    try:
        assert len(doc.tables) == 1
        tbl = doc.tables[0]

        rng = tbl.get_table_range()
        assert rng is not None

        assert len(rng) == 100

        i = 0
        for cell in tbl:
            assert cell is not None
            i += 1
        assert i == 100

    finally:
        doc.close()


def test_insert_cols_rows(loader):
    from ooodev.utils.table_helper import TableHelper

    doc = WriteDoc.create_doc(loader)
    try:
        tbl_data = TableHelper.make_2d_array(num_rows=5, num_cols=5)

        cursor = doc.get_cursor()

        tbl = cursor.add_table(table_data=tbl_data)
        assert len(tbl.rows) == 5
        assert len(tbl.columns) == 5
        tbl.rows.append_rows(2)
        assert len(tbl.rows) == 7
        tbl.columns.append_columns(2)
        assert len(tbl.columns) == 7

        rng = tbl.get_table_range()
        assert rng.col_count == 7
        assert rng.row_count == 7

        tbl.columns.remove_by_index(-1)
        del tbl.columns[-1]

        tbl.rows.remove_by_index(-1)
        del tbl.rows[-1]

        assert len(tbl.rows) == 5
        assert len(tbl.columns) == 5

        assert len(doc.tables) == 1

        tbl.ensure_colum_row("J10")
        assert len(tbl.rows) == 10
        assert len(tbl.columns) == 10

    finally:
        doc.close()


def test_table_cursor(loader):
    from ooodev.utils.table_helper import TableHelper

    doc = WriteDoc.create_doc(loader)
    try:
        tbl_data = TableHelper.make_2d_array(num_rows=10, num_cols=10)

        cursor = doc.get_cursor()

        tbl = cursor.add_table(table_data=tbl_data)
        assert len(tbl.rows) == 10
        assert len(tbl.columns) == 10
        cursor = tbl.create_cursor_by_cell_name("A1")
        assert cursor is not None
        assert cursor.get_range_name() == "A1"
        rng = cursor.get_range_obj()
        assert str(rng) == "A1:A1"
        cursor.go_right(2, True)
        assert cursor.get_range_name() == "A1:C1"
        rng = cursor.get_range_obj()
        assert str(rng) == "A1:C1"
        cursor.go_down(2, True)
        assert cursor.get_range_name() == "A1:C3"
        rng = cursor.get_range_obj()
        assert str(rng) == "A1:C3"
        cursor.goto_end(True)
        assert cursor.get_range_name() == "A1:J10"
        rng = cursor.get_range_obj()
        assert str(rng) == "A1:J10"
        cursor.goto_start()
        assert cursor.get_range_name() == "A1"
        rng = cursor.get_range_obj()
        assert str(rng) == "A1:A1"
        cursor.goto_end()
        assert cursor.get_range_name() == "J10"
        rng = cursor.get_range_obj()
        assert str(rng) == "J10:J10"

    finally:
        doc.close()


def test_column_separators(loader):
    from ooodev.utils.table_helper import TableHelper

    doc = WriteDoc.create_doc(loader)
    try:
        # Moving column separators is how column widths are adjusted.
        tbl_data = TableHelper.make_2d_array(num_rows=4, num_cols=4)

        cursor = doc.get_cursor()

        _ = cursor.add_table(table_data=tbl_data)
        tbl = doc.tables[0]

        assert len(tbl.table_column_separators) == 3
        sep1 = tbl.table_column_separators[0]
        tbl.table_column_separators[0].position += 10
        assert sep1.position < tbl.table_column_separators[0].position

    finally:
        doc.close()


def test_set_table_array(loader, bond_movies_table: list):

    doc = WriteDoc.create_doc(loader)
    try:
        tbl_data = bond_movies_table  # TableHelper.make_2d_array(num_rows=4, num_cols=4)

        cursor = doc.get_cursor()

        _ = cursor.add_table(table_data=[[[]]], first_row_header=False)
        tbl = doc.tables[0]
        rng = tbl.range_converter.get_range_from_2d(tbl_data)

        # tbl.ensure_colum_row(col=3, row=24)
        tbl.ensure_colum_row(rng.cell_end)
        assert len(tbl.rows) == 25
        assert len(tbl.columns) == 4
        # when setting table data the table must be the same size as the data
        tbl.set_data_array(tbl_data)

        cell = tbl["D23"]
        assert cell.value == "Marc Forster"

    finally:
        doc.close()


def test_set_array_by_cell(loader, bond_movies_table: list):

    doc = WriteDoc.create_doc(loader)
    try:
        tbl_data = bond_movies_table  # TableHelper.make_2d_array(num_rows=4, num_cols=4)

        cursor = doc.get_cursor()

        _ = cursor.add_table(table_data=[[[]]], first_row_header=False)
        tbl = doc.tables[0]
        rng = tbl.range_converter.get_range_from_2d(tbl_data)
        rng_addr = rng.get_cell_range_address()
        assert rng_addr.EndRow == 24  # zero based
        assert rng_addr.EndColumn == 3  # zero based

        tbl.set_data_array_cell(tbl_data, cell="B2", ensure_cols=True, ensure_rows=True)

        assert len(tbl.rows) == 26
        assert len(tbl.columns) == 5

        cell = tbl["E24"]
        assert cell.value == "Marc Forster"

    finally:
        doc.close()


def test_range_sub_range(loader, copy_fix_writer):
    fnm = copy_fix_writer("bondMovies.odt")
    doc = WriteDoc.open_doc(fnm=fnm, loader=loader)

    try:
        tbl = doc.tables[0]
        cell_rng = tbl.get_cell_range_by_name("A1:D10")
        cell = cell_rng["A2"]
        assert cell.cell_name == "A2"
        assert cell.value == "Dr. No"

        sub_rng = cell_rng.get_cell_range("B2:D4")
        # relative, C# is actually D4
        sub_cell = sub_rng["c3"]
        assert sub_cell.cell_name == "D4"
        assert sub_cell.value == "Guy Hamilton"

    finally:
        doc.close()
