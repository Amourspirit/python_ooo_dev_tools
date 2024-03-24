from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_add_button(loader) -> None:
    # Test adding controls to a cell and a range
    # The controls are found when calling "cell.control.current_control" by shape position.
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.sheets[0]

        cell = sheet["A1"]
        chk = cell.control.insert_control_check_box("My CheckBox", tri_state=False)
        assert chk is not None

        cell = sheet["A1"]
        chk = cell.control.current_control
        assert chk is not None

        cell = sheet["A2"]
        combo = cell.control.insert_control_combo_box(entries=["A", "B", "C"])
        assert combo is not None

        cell = sheet["A2"]
        combo = cell.control.current_control
        assert combo is not None

        cell = sheet["A3"]
        currency = cell.control.insert_control_currency_field(spin_button=True)
        assert currency is not None

        cell = sheet["A3"]
        currency = cell.control.current_control
        assert currency is not None

        cell = sheet["A4"]
        date = cell.control.insert_control_date_field()
        assert date is not None

        cell = sheet["A4"]
        date = cell.control.current_control
        assert date is not None

        cell = sheet["A5"]
        file = cell.control.insert_control_file()
        assert file is not None

        cell = sheet["A5"]
        file = cell.control.current_control
        assert file is not None

        cell = sheet["A6"]
        ff = cell.control.insert_control_formatted_field(spin_button=True)
        assert ff is not None

        cell = sheet["A6"]
        ff = cell.control.current_control
        assert ff is not None

        cell = sheet["A7"]
        img_btn = cell.control.insert_control_image_button()
        assert img_btn is not None

        cell = sheet["A7"]
        img_btn = cell.control.current_control
        assert img_btn is not None

        cell = sheet["A8"]
        lbl = cell.control.insert_control_label("My Label")
        assert lbl is not None

        cell = sheet["A8"]
        lbl = cell.control.current_control
        assert lbl is not None

        cell = sheet["A9"]
        num = cell.control.insert_control_numeric_field(spin_button=True)
        assert num is not None

        cell = sheet["A9"]
        num = cell.control.current_control
        assert num is not None

        cell = sheet["A10"]
        pattern = cell.control.insert_control_pattern_field()
        assert pattern is not None

        cell = sheet["A10"]
        pattern = cell.control.current_control
        assert pattern is not None

        cell = sheet["A11"]
        rb = cell.control.insert_control_radio_button("My Radio Button")
        assert rb is not None

        cell = sheet["A11"]
        rb = cell.control.current_control
        assert rb is not None

        cell = sheet["A12"]
        spin = cell.control.insert_control_spin_button()
        assert spin is not None

        cell = sheet["A12"]
        spin = cell.control.current_control
        assert spin is not None

        cell = sheet["A13"]
        text_field = cell.control.insert_control_text_field("Hello World")
        assert text_field is not None

        cell = sheet["A13"]
        text_field = cell.control.current_control
        assert text_field is not None

        cell = sheet["A14"]
        time = cell.control.insert_control_time_field()
        assert time is not None

        cell = sheet["A14"]
        time = cell.control.current_control
        assert time is not None

        cell = sheet["B3"]
        btn = cell.control.insert_control_button("My Button")
        assert btn is not None

        cell = sheet["B3"]
        btn = cell.control.current_control

        rng = sheet.get_range(range_name="C3:E5")
        rng_btn = rng.control.insert_control_button("My Rng Button")
        assert rng_btn is not None

        rng = sheet.get_range(range_name="B3:D5")
        rng_btn = rng.control.current_control
        assert rng_btn is not None

        rng = sheet.get_range(range_name="B6:D9")
        gb = rng.control.insert_control_group_box("My Group Box")
        assert gb is not None

        rng = sheet.get_range(range_name="B6:D9")
        gb = rng.control.current_control
        assert gb is not None

        rng = sheet.get_range(range_name="f3:h5")
        grid = rng.control.insert_control_grid("My Grid")
        assert grid is not None

        rng = sheet.get_range(range_name="f3:h5")
        grid = rng.control.current_control
        assert grid is not None

        rng = sheet.get_range(range_name="b10:c12")
        lb = rng.control.insert_control_list_box(entries=["D", "E", "F"], drop_down=False)
        assert lb is not None

        rng = sheet.get_range(range_name="b10:c13")
        lb = rng.control.current_control
        assert lb is not None

        rng = sheet.get_range(range_name="b14:c18")
        rich_text = rng.control.insert_control_rich_text()
        rich_text.model.MultiLine = True
        assert rich_text is not None

        rng = sheet.get_range(range_name="b14:c18")
        rich_text = rng.control.current_control
        assert rich_text is not None

        rng = sheet.get_range(range_name="b19:e19")
        scroll = rng.control.insert_control_scroll_bar()
        assert scroll is not None

        rng = sheet.get_range(range_name="b19:e19")
        scroll = rng.control.current_control
        assert scroll is not None

    finally:
        doc.close()
