from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.draw import Draw, DrawDoc, ImpressDoc
from ooodev.utils.lo import Lo


def test_master_page_draw(loader) -> None:
    doc = DrawDoc(Draw.create_draw_doc(loader))
    assert isinstance(doc, DrawDoc)
    slide = doc.get_slide(idx=0)
    assert slide is not None

    page = doc.insert_master_page(idx=0)
    assert page is not None
    doc.close_doc()  # type: ignore


def test_master_page_impress(loader) -> None:
    doc = ImpressDoc(Draw.create_impress_doc(loader))
    assert isinstance(doc, ImpressDoc)

    slide = doc.get_slide(idx=0)
    assert slide is not None

    page = doc.insert_master_page(idx=0)
    assert page is not None

    doc.close_doc()


def test_get_handout_master_page(loader) -> None:
    doc = ImpressDoc(Draw.create_impress_doc(loader))
    slide = doc.get_slide(idx=0)
    assert slide is not None

    page = doc.get_handout_master_page()
    assert page is not None
    assert page.owner is doc

    doc.close_doc()


def test_rectangle(loader) -> None:
    x = 10
    y = 20
    width = 17
    height = 12
    doc = DrawDoc(Draw.create_draw_doc(loader))
    slide = doc.get_slide(idx=0)

    num_shapes = slide.get_count()
    assert num_shapes == 0
    shape = slide.draw_rectangle(x=x, y=y, width=width, height=height)
    pos1 = shape.get_position()
    assert pos1.X == x * 100
    assert pos1.Y == y * 100
    pos2 = shape.get_position_mm()
    assert pos2.X == x
    assert pos2.Y == y
    size1 = shape.get_size()
    assert size1.Height == height * 100
    assert size1.Width == width * 100
    size2 = shape.get_size_mm()
    assert size2.Height == height
    assert size2.Width == width

    assert shape is not None

    num_shapes2 = slide.get_count()
    assert num_shapes2 == 1
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y
    size2 = t_shape.get_size()
    assert size2.Width == size1.Width
    assert size2.Height == size1.Height

    doc.close_doc()


def test_circle(loader) -> None:
    x = 10
    y = 20
    radius = 20
    doc = DrawDoc(Draw.create_draw_doc(loader))
    slide = doc.get_slide(idx=0)

    num_shapes = slide.get_count()
    assert num_shapes == 0
    shape = slide.draw_circle(x=x, y=y, radius=radius)
    pos1 = shape.get_position()

    num_shapes2 = slide.get_count()
    assert num_shapes2 == 1
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    doc.close_doc()


def test_ellipse(loader) -> None:
    x = 100
    y = 100
    width = 75
    height = 25
    doc = DrawDoc(Draw.create_draw_doc(loader))
    slide = doc.get_slide(idx=0)

    num_shapes = slide.get_count()
    assert num_shapes == 0
    shape = slide.draw_ellipse(x=x, y=y, width=width, height=height)
    pos1 = shape.get_position()

    num_shapes2 = slide.get_count()
    assert num_shapes2 == 1
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    doc.close_doc()


def test_text(loader) -> None:
    x = 120
    y = 120
    width = 60
    height = 30
    doc = DrawDoc(Draw.create_draw_doc(loader))
    slide = doc.get_slide(idx=0)

    num_shapes = slide.get_count()
    assert num_shapes == 0
    shape = slide.draw_text(msg="Hello LibreOffice", x=x, y=y, width=width, height=height, font_size=24)
    pos1 = shape.get_position()

    num_shapes2 = slide.get_count()
    assert num_shapes2 == 1
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    doc.close_doc()


def test_line(loader) -> None:
    doc = DrawDoc(Draw.create_draw_doc(loader))
    slide = doc.get_slide(idx=0)

    num_shapes = slide.get_count()
    assert num_shapes == 0
    shape = slide.draw_line(x1=50, y1=50, x2=200, y2=200)
    pos1 = shape.get_position()

    num_shapes2 = slide.get_count()
    assert num_shapes2 == 1
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    doc.close_doc()


def test_polar_line(loader) -> None:
    doc = DrawDoc(Draw.create_draw_doc(loader))
    slide = doc.get_slide(idx=0)

    num_shapes = slide.get_count()
    assert num_shapes == 0
    shape = slide.draw_polar_line(x=60, y=200, degrees=45, distance=100)
    pos1 = shape.get_position()

    num_shapes2 = slide.get_count()
    assert num_shapes2 == 1
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    doc.close_doc()


def test_lines_and_move_to_bottom(loader) -> None:
    doc = DrawDoc(Draw.create_draw_doc(loader))
    from ooodev.units import UnitMM

    slide = doc.get_slide(idx=0)

    num_shapes = slide.get_count()
    assert num_shapes == 0
    line = slide.draw_line(x1=UnitMM(50), y1=UnitMM(50), x2=UnitMM(200), y2=UnitMM(200))
    polar_line = slide.draw_polar_line(x=60, y=200, degrees=45, distance=100)
    line_pos = line.get_position()
    polar_line_pos = polar_line.get_position()

    num_shapes2 = slide.get_count()
    assert num_shapes2 == 2
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == polar_line_pos.X
    assert pos2.Y == polar_line_pos.Y

    polar_line.move_to_bottom()
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == line_pos.X
    assert pos2.Y == line_pos.Y

    doc.close_doc()


def test_lines_and_move_to_top(loader) -> None:
    from ooodev.units import UnitMM

    doc = DrawDoc(Draw.create_draw_doc(loader))
    slide = doc.get_slide(idx=0)

    num_shapes = slide.get_count()
    assert num_shapes == 0
    line = slide.draw_line(x1=50, y1=50, x2=200, y2=200)
    polar_line = slide.draw_polar_line(x=UnitMM(60), y=UnitMM(200), degrees=45, distance=UnitMM(100))
    line_pos = line.get_position()
    polar_line_pos = polar_line.get_position()

    num_shapes2 = slide.get_count()
    assert num_shapes2 == 2
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == polar_line_pos.X
    assert pos2.Y == polar_line_pos.Y

    line.move_to_top()
    t_shape = slide.find_top_shape()
    pos2 = t_shape.get_position()
    assert pos2.X == line_pos.X
    assert pos2.Y == line_pos.Y

    doc.close_doc()


def test_cursor(loader) -> None:
    doc = DrawDoc(Draw.create_draw_doc(loader))
    slide = doc.get_slide(idx=0)

    rect = slide.draw_rectangle(x=10, y=10, width=10, height=10)
    cursor = rect.get_shape_text_cursor()
    assert cursor is not None
    cursor.append_para("Hello World")
    cursor.goto_start()
    cursor.goto_end(True)
    assert cursor.get_string().startswith("Hello World")

    doc.close_doc()
