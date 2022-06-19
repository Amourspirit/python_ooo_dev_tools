import os
import pytest
from pathlib import Path
from typing import cast

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_build_doc(loader, props_str_to_dict, fix_image_path, capsys: pytest.CaptureFixture):
    from ooodev.utils.lo import Lo
    from ooodev.office.write import Write
    from ooodev.utils.gui import GUI
    from ooodev.utils.props import Props
    from ooodev.utils.color import CommonColor
    from ooodev.utils.images import Images
    from ooodev.utils.info import Info
    from ooodev.exceptions import ex
    from functools import partial

    visible = True
    delay = 2000
    lock = False
    doc = Write.create_doc(loader)
    try:
        if lock:
            Lo.lock_controllers()
        if visible:
            GUI.set_visible(visible, doc)

        cursor = Write.get_cursor(doc)
        append = partial(Write.append, cursor)
        para = partial(Write.append_para, cursor)
        nl = partial(Write.append_line, cursor)
        get_pos = partial(Write.get_position, cursor)

        im_fnm = cast(Path, fix_image_path("skinner.png"))

        capsys.readouterr()  # clear buffer
        Props.show_obj_props(prop_kind="Cursor", obj=cursor)
        cap = capsys.readouterr()
        cap_out = cap.out
        assert cap_out is not None
        prop_dict: dict = props_str_to_dict(cap_out)
        assert prop_dict["ParaBackTransparent"] == "True"
        assert prop_dict["Endnote"] == "None"
        assert prop_dict["Footnote"] == "None"
        assert prop_dict["TextField"] == "None"
        assert prop_dict["TextTable"] == "None"

        append(text="Some examples of simple text ")
        pos = get_pos()
        assert pos == 29
        append("styles.")
        append(ctl_char=Write.ControlCharacter.LINE_BREAK)
        Write.style_left_bold(cursor=cursor, pos=pos)

        pos = get_pos()
        assert pos == 37
        para("This line is written in red italics.")
        Write.style_left_color(cursor=cursor, pos=pos, color=CommonColor.DARK_RED)
        Write.style_left_italic(cursor=cursor, pos=pos)

        Write.append_para(cursor=cursor, text="Back to old style")
        nl()

        Write.append_para(cursor=cursor, text="A Nice Big Heading")
        Write.style_prev_paragraph(cursor, "Heading 1")

        Write.append_para(cursor, "The following points are important:")
        pos = get_pos()
        assert pos == 148
        Write.append_para(cursor, "Have a good breakfast")
        Write.append_para(cursor, "Have a good lunch")
        with pytest.raises(ex.PropertyNotFoundError):
            Write.style_prev_paragraph(cursor, "NumberingStyleName", "Numbering 1")
        Write.append_para(cursor, "Have a good dinner")
        nl()
        tvc = Write.get_view_cursor(doc)
        tvc.gotoEnd(False)
        # tvc.goRight(pos, False)
        # cpos = Write.get_position(cursor)
        # cp = cpos-pos
        # cursor.goLeft(cp, True)
        # cursor.NumberingStyleName = "Numbering 1"
        # cursor.num
        # cursor.goRight(cp, False)
        # cursor.NumberingStyleName = ""
        Write.style_left(tvc, pos, "NumberingStyleName", "Numbering 1")
        # numberingLevel
        # Write.style_left(cursor, pos, "NumberingIsNumber", True)
        # Write.style_left(cursor, pos, "NumberingLevel", 1)
        # "23250040901"
        # https://www.openoffice.org/api/docs/common/ref/com/sun/star/style/ParagraphProperties.html#NumberingStyleName

        para("Breakfast should include:")
        para("Porridge")
        pos = get_pos()
        para("Orange Juice")
        para("A Cup of Tea")
        nl()
        
        Write.style_left(cursor, pos, "NumberingStyleName", "Numbering 1")
        Write.style_left(cursor, pos, "ParaIsNumberingRestart", True)

        append("This line ends with a bookmark.")
        Write.add_bookmark(cursor=cursor, name="ad-bookmark")
        para("\n")

        para("Here's some code:")

        tvc = Write.get_view_cursor(doc)
        tvc.gotoRange(cursor.getEnd(), False)
        # tvc.gotoEnd(False)
        coord = Write.get_coord_str(tvc)
        # assert coord == "(25285, 9398)" # "(22412, 9885)"
        ypos = tvc.getPosition().Y

        nl()
        nl("public class Hello")
        pos = get_pos()
        nl("{")
        nl("  public static void main(String args[]")
        nl('  {  System.out.println("Hello World");  }')
        para("}  // end of Hello class")
        nl()
        Write.style_left_code(cursor, pos)

        para("A text frame")
        # page_cursor = Write.get_page_cursor(doc)
        # pg = page_cursor.getPage()
        pg = Write.get_current_page(tvc)
        Write.add_text_frame(
            cursor=cursor,
            ypos=ypos,
            text="This is a newly created text frame.\nWhich is over on the right of the page, next to the code.",
            page_num=pg,
            width=4000,
            height=1500,
        )

        para("A link to my JLOP website:")
        url_str = "http://fivedots.coe.psu.ac.th/~ad/jlop/"
        pos = get_pos()
        append(url_str)

        Write.style_left(cursor=cursor, pos=pos, prop_name="HyperLinkURL", prop_val=url_str)
        Write.end_paragraph(cursor)

        Write.page_break(cursor)

        with Lo.ControllerLock():
            Lo.delay(delay)
            para("Image Example")
            Write.style_prev_paragraph(cursor, "Heading 2")

            para(f"The following image comes from{im_fnm.name}")
            nl()
            
            append(f"Image Link is {im_fnm.name}: ")

            img_size = Images.get_size_100mm(im_fnm=im_fnm)
            assert img_size.Height == 5751
            assert img_size.Width == 6092
            # Write.add_image_link(doc, cursor, im_fnm, img_size.Width, img_size.Height)
            Write.add_image_link(doc, cursor, im_fnm)

            # enlarge by 1.5x
            h = round(img_size.Height * 1.5)
            w = round(img_size.Width * 1.5)

            Write.add_image_link(doc, cursor, im_fnm, w, h)

        Write.end_paragraph(cursor)
        Lo.delay(delay)

        Write.style_prev_paragraph(cursor=cursor, prop_name="ParaAdjust", prop_val=Write.ParagraphAdjust.CENTER)

        append("Image as a shape: ")
        Write.add_image_shape(doc=doc, cursor=cursor, fnm=im_fnm)
        Write.end_paragraph(cursor)
        Lo.delay(delay)

        text_width = Write.get_page_text_width(doc)
        assert text_width > 0
        Write.add_line_divider(cursor=cursor, line_width=round(text_width * 0.5))

        Info.set_doc_props(doc=doc, subject="Writer Text Example", title="Examples", author=":Barry-Thomas-Paul: Moss")
        Lo.delay(delay)

        # move view cursor to bookmark position
        bookmark = Write.find_bookmark(doc, "ad-bookmark")
        bm_range = bookmark.getAnchor()

        view_cursor = Write.get_view_cursor(doc)
        view_cursor.gotoRange(bm_range, False)
        if lock:
            Lo.unlock_controllers()
        Lo.delay(delay)
        
    finally:

        Lo.close_doc(doc, False)
