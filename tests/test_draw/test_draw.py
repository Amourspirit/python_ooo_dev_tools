from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.office.draw import Draw
from ooodev.loader.lo import Lo


def test_master_page_draw(loader) -> None:
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)
    assert slide is not None

    page = Draw.insert_master_page(doc=doc, idx=0)
    assert page is not None

    Lo.close(doc)  # type: ignore


def test_master_page_impress(loader) -> None:
    doc = Draw.create_impress_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)
    assert slide is not None

    page = Draw.insert_master_page(doc=doc, idx=0)
    assert page is not None

    Lo.close(doc)  # type: ignore


def test_get_handout_master_page(loader) -> None:
    doc = Draw.create_impress_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)
    assert slide is not None

    page = Draw.get_handout_master_page(doc=doc)
    assert page is not None

    Lo.close(doc)  # type: ignore


def test_rectangle(loader) -> None:
    x = 10
    y = 20
    width = 17
    height = 12
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    shape = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
    pos1 = shape.getPosition()
    assert pos1.X == x * 100
    assert pos1.Y == y * 100
    size1 = shape.getSize()
    assert size1.Height == height * 100
    assert size1.Width == width * 100
    assert shape is not None

    num_shapes2 = slide.getCount()
    assert num_shapes2 == 1
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y
    size2 = t_shape.getSize()
    assert size2.Width == size1.Width
    assert size2.Height == size1.Height

    Lo.close(doc)  # type: ignore


def test_rectangle_no_loader(loader) -> None:
    x = 10
    y = 20
    width = 17
    height = 12
    doc = Draw.create_draw_doc()
    # defaults to index 0
    slide = Draw.get_slide(doc=doc)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    shape = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
    pos1 = shape.getPosition()
    assert pos1.X == x * 100
    assert pos1.Y == y * 100
    size1 = shape.getSize()
    assert size1.Height == height * 100
    assert size1.Width == width * 100
    assert shape is not None

    num_shapes2 = slide.getCount()
    assert num_shapes2 == 1
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y
    size2 = t_shape.getSize()
    assert size2.Width == size1.Width
    assert size2.Height == size1.Height

    Lo.close(closeable=doc, deliver_ownership=False)


def test_circle(loader) -> None:
    x = 10
    y = 20
    radius = 20
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    shape = Draw.draw_circle(slide=slide, x=x, y=y, radius=radius)
    pos1 = shape.getPosition()

    num_shapes2 = slide.getCount()
    assert num_shapes2 == 1
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    Lo.close(doc)  # type: ignore


def test_ellipse(loader) -> None:
    x = 100
    y = 100
    width = 75
    height = 25
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    shape = Draw.draw_ellipse(slide=slide, x=x, y=y, width=width, height=height)
    pos1 = shape.getPosition()

    num_shapes2 = slide.getCount()
    assert num_shapes2 == 1
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    Lo.close(doc)  # type: ignore


def test_text(loader) -> None:
    x = 120
    y = 120
    width = 60
    height = 30
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    shape = Draw.draw_text(slide=slide, msg="Hello LibreOffice", x=x, y=y, width=width, height=height, font_size=24)
    pos1 = shape.getPosition()

    num_shapes2 = slide.getCount()
    assert num_shapes2 == 1
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    Lo.close(doc)  # type: ignore


def test_line(loader) -> None:
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    shape = Draw.draw_line(slide=slide, x1=50, y1=50, x2=200, y2=200)
    pos1 = shape.getPosition()

    num_shapes2 = slide.getCount()
    assert num_shapes2 == 1
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    Lo.close(doc)  # type: ignore


def test_polar_line(loader) -> None:
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    shape = Draw.draw_polar_line(slide=slide, x=60, y=200, degrees=45, distance=100)
    pos1 = shape.getPosition()

    num_shapes2 = slide.getCount()
    assert num_shapes2 == 1
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == pos1.X
    assert pos2.Y == pos1.Y

    Lo.close(doc)  # type: ignore


def test_lines_and_move_to_bottom(loader) -> None:
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    line = Draw.draw_line(slide=slide, x1=50, y1=50, x2=200, y2=200)
    polar_line = Draw.draw_polar_line(slide=slide, x=60, y=200, degrees=45, distance=100)
    line_pos = line.getPosition()
    polar_line_pos = polar_line.getPosition()

    num_shapes2 = slide.getCount()
    assert num_shapes2 == 2
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == polar_line_pos.X
    assert pos2.Y == polar_line_pos.Y

    Draw.move_to_bottom(slide=slide, shape=polar_line)
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == line_pos.X
    assert pos2.Y == line_pos.Y

    Lo.close(doc)  # type: ignore


def test_lines_and_move_to_top(loader) -> None:
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    line = Draw.draw_line(slide=slide, x1=50, y1=50, x2=200, y2=200)
    polar_line = Draw.draw_polar_line(slide=slide, x=60, y=200, degrees=45, distance=100)
    line_pos = line.getPosition()
    polar_line_pos = polar_line.getPosition()

    num_shapes2 = slide.getCount()
    assert num_shapes2 == 2
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == polar_line_pos.X
    assert pos2.Y == polar_line_pos.Y

    Draw.move_to_top(slide=slide, shape=line)
    t_shape = Draw.find_top_shape(slide)
    pos2 = t_shape.getPosition()
    assert pos2.X == line_pos.X
    assert pos2.Y == line_pos.Y

    Lo.close(doc)  # type: ignore


def _test_shape_drawing_object(loader) -> None:
    from ooodev.utils.gui import GUI

    # With the following
    # nothing appears on the slide, but object is there
    # position, width and height can not be changed

    # com.sun.star.drawing.Shape3DSceneObject # nothing show up on screen but test pass
    # com.sun.star.drawing.Shape3DCubeObject # nothing show up on screen but test pass
    # com.sun.star.drawing.Shape3DSphereObject # nothing show up on screen but test pass
    # com.sun.star.drawing.Shape3DLatheObject # nothing show up on screen but test pass
    # com.sun.star.drawing.Shape3DExtrudeObject # nothing show up on screen but test pass
    # com.sun.star.drawing.Shape3DPolygonObject # nothing show up on screen but test pass
    doc = Draw.create_draw_doc(loader)
    slide = Draw.get_slide(doc=doc, idx=0)

    GUI.set_visible(is_visible=True, odoc=doc)
    Lo.delay(1000)
    GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

    num_shapes = slide.getCount()
    assert num_shapes == 0
    Draw.add_shape(slide=slide, shape_type="Shape3DPolygonObject", x=120, y=120, width=60, height=60)
    num_shapes2 = slide.getCount()
    assert num_shapes2 == 1

    Lo.delay(3000)
    Lo.close(doc)  # type: ignore
