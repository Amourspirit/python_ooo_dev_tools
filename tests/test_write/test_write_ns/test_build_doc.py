import os
import pytest
from pathlib import Path
from typing import cast

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.lo import Lo
from ooodev.write import Write, ControlCharacterEnum, ParagraphAdjust
from ooodev.write import WriteDoc
from ooodev.utils.props import Props
from ooodev.utils.color import CommonColor
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.info import Info
from ooodev.exceptions import ex
from ooodev.utils.date_time_util import DateUtil


@pytest.mark.skip_not_headless_os("linux", "Errors When GUI is present")
def test_build_doc(loader, props_str_to_dict, fix_image_path, capsys: pytest.CaptureFixture):
    # see comments Write.add_image_shape(cursor=cursor, fnm=im_fnm) Line: 181

    visible = False if Lo.bridge_connector.headless else True
    delay = 0  # 500
    doc = WriteDoc.create_doc(loader=loader, visible=visible)
    try:
        cursor = doc.get_cursor()

        im_fnm = cast(Path, fix_image_path("skinner.png"))

        capsys.readouterr()  # clear buffer
        Props.show_obj_props(prop_kind="Cursor", obj=cursor.component)
        cap = capsys.readouterr()
        cap_out = cap.out
        assert cap_out is not None
        prop_dict: dict = props_str_to_dict(cap_out)
        assert prop_dict["ParaBackTransparent"] == "True"
        assert prop_dict["ParaBackColor"] == "-1"
        assert prop_dict["Endnote"] == "None"
        assert prop_dict["Footnote"] == "None"
        assert prop_dict["TextField"] == "None"
        assert prop_dict["TextTable"] == "None"

        cursor.append(text="Some examples of simple text ")
        pos = cursor.get_pos()
        assert pos == 29
        cursor.append("styles.")
        cursor.append(ctl_char=ControlCharacterEnum.LINE_BREAK)
        cursor.style_left_bold(pos=pos)

        pos = cursor.get_pos()
        assert pos == 37
        cursor.append_para("This line is written in red italics.")
        cursor.style_left_color(pos=pos, color=CommonColor.DARK_RED)
        cursor.style_left_italic(pos=pos)

        cursor.append_para(text="Back to old style")
        cursor.append_line()

        cursor.append_para(text="A Nice Big Heading")
        cursor.style_prev_paragraph("Heading 1")

        cursor.append_para("The following points are important:")
        pos = cursor.get_pos()
        assert pos == 148
        cursor.append_para("Have a good breakfast")
        cursor.append_para("Have a good lunch")
        with pytest.raises(ex.PropertyNotFoundError):
            cursor.style_prev_paragraph("NumberingStyleName", "Numbering 1")
        cursor.append_para("Have a good dinner")

        # Write.style_left(cursor, pos, "NumberingStyleName", "Numbering 1")
        cursor.style_left(pos, "NumberingStyleName", "Numbering 123")
        # Numbering 123

        tvc = doc.get_view_cursor()
        # tvc.gotoEnd(False)
        # Write.dispatch_cmd_left(vcursor=tvc, pos=pos, cmd="DefaultNumbering", toggle=True)
        tvc.append_para()
        # https://www.openoffice.org/api/docs/common/ref/com/sun/star/style/ParagraphProperties.html#NumberingStyleName

        tvc.append_para("Breakfast should include:")
        pos = tvc.get_pos()
        tvc.append_para("Porridge")
        tvc.append_para("Orange Juice")
        tvc.append_para("A Cup of Tea")
        tvc.style_left(pos, "NumberingStyleName", "Numbering abc")
        # tvc.gotoEnd(False)
        # Write.dispatch_cmd_left(vcursor=tvc, pos=pos, cmd="DefaultNumbering", toggle=True)
        tvc.append_para()

        tvc.append("This line ends with a bookmark.")
        tvc.add_bookmark(name="ad-bookmark")
        tvc.append_para("\n")

        tvc.append_para("Here's some code:")

        tvc = doc.get_view_cursor()
        tvc.component.gotoRange(cursor.component.getEnd(), False)
        # tvc.gotoEnd(False)
        coord = tvc.get_coord_str()
        # assert coord == "(25285, 9398)" # "(22412, 9885)"
        ypos = tvc.get_position().Y

        tvc.append_para()
        pos = tvc.get_pos()
        tvc.append_line("public class Hello")
        tvc.append_line("{")
        tvc.append_line("  public static void main(String args[]")
        tvc.append_line('  {  System.out.println("Hello World");  }')
        tvc.append_para("}  // end of Hello class")
        tvc.append_line()
        tvc.style_left_code(pos)
        cursor.goto_end()
        cursor.style_prev_paragraph(prop_name="ParaBackColor", prop_val=CommonColor.LIGHT_GRAY)

        tvc.append_para("A text frame")
        # page_cursor = Write.get_page_cursor(doc)
        # pg = page_cursor.getPage()
        pg = tvc.get_current_page()
        frame = tvc.add_text_frame(
            ypos=ypos,
            text="This is a newly created text frame.\nWhich is over on the right of the page, next to the code.",
            page_num=pg,
            width=4000,
            height=1500,
        )
        assert frame is not None

        tvc.append("A link to ")
        pos = tvc.get_pos()
        tvc.append("OOO Development Tools")
        url_str = "https://github.com/Amourspirit/python_ooo_dev_tools"
        tvc.style_left(pos=pos, prop_name="HyperLinkURL", prop_val=url_str)
        tvc.append(" Website.")
        tvc.end_paragraph()

        tvc.page_break()

        with Lo.ControllerLock():
            Lo.delay(delay)
            cursor.goto_end()
            ("Image Example")
            cursor.style_prev_paragraph("Heading 2")

            cursor.append_para(f'The following image comes from "{im_fnm.name}":')
            cursor.append_para()

            cursor.append("Image as a link: ")

            img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)
            assert img_size.Height == 5751
            assert img_size.Width == 6092
            # Write.add_image_link(doc, cursor, im_fnm, img_size.Width, img_size.Height)
            cursor.add_image_link(fnm=im_fnm, width=img_size.Width, height=img_size.Height)

            # enlarge by 1.5x
            h = round(img_size.Height * 1.5)
            w = round(img_size.Width * 1.5)

            cursor.add_image_link(fnm=im_fnm, width=w, height=h)
            cursor.end_paragraph()

        Lo.delay(delay)

        cursor.style_prev_paragraph(prop_name="ParaAdjust", prop_val=ParagraphAdjust.CENTER)

        # append("Image as a shape: ")
        cursor.append_line(text="Image as Shape:")

        # for some unknown reason when image shape is added in linux in GUI mode test will fail drastically.
        #   tests/test_write/test_build_doc.py terminate called after throwing an instance of 'com::sun::star::lang::DisposedException'
        #   Fatal Python error: Aborted
        # on windows is fine. Running on linux in headless fine.
        _ = cursor.add_image_shape(fnm=im_fnm)

        cursor.end_paragraph()
        Lo.delay(delay)

        text_width = doc.get_page_text_width()
        assert text_width > 0
        cursor.add_line_divider(line_width=round(text_width * 0.5))

        cursor.append_para("\nTimestamp: " + DateUtil.time_stamp() + "\n")
        cursor.append("Time (according to office): ")
        cursor.append_date_time()
        cursor.end_paragraph()

        Info.set_doc_props(
            doc=doc.component, subject="Writer Text Example", title="Examples", author=":Barry-Thomas-Paul: Moss"
        )
        Lo.delay(delay)

        # move view cursor to bookmark position
        bookmark = doc.find_bookmark("ad-bookmark")
        assert bookmark is not None
        bm_range = bookmark.get_anchor()

        view_cursor = doc.get_view_cursor()
        view_cursor.goto_range(bm_range, False)

        # GUI.set_visible(True, doc)
        Lo.delay(delay)

    finally:
        doc.close_doc()
