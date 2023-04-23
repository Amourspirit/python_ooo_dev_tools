.. _help_writer_format_style_para:

Write Style Para Class
======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.writer.style.Para` class is used to set the paragraph style.

:ref:`ch06` has more information on paragraph styles.

Setup
-----

General function used to run these examples.

.. must be before the tabs directive
.. include:: ../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            import uno
            from com.sun.star.text import XTextDocument

            from ooodev.format.writer.direct.char.font import Font
            from ooodev.format.writer.direct.para.area import Color as ParaBgColor
            from ooodev.format.writer.direct.para.indent_space import Spacing, LineSpacing, ModeKind
            from ooodev.format.writer.style import Para as StylePara
            from ooodev.office.write import Write
            from ooodev.units import UnitMM
            from ooodev.utils.color import StandardColor
            from ooodev.utils.gui import GUI
            from ooodev.utils.lo import Lo


            def main() -> int:
                p_txt = (
                    |short_ptext|
                )

                with Lo.Loader(Lo.ConnectPipe()):
                    doc = Write.create_doc()
                    GUI.set_visible(doc)
                    Lo.delay(300)
                    GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
                    new_style_name = "CoolParagraph"
                    if not create_para_style(doc, new_style_name):
                        raise RuntimeError(f"Could not create new paragraph style: {new_style_name}")

                    # get the XTextRange of the document
                    xtext_range = doc.getText().getStart()
                    # Load the paragraph style and apply it to the text range.
                    para_style = StylePara(new_style_name)
                    para_style.apply(xtext_range)

                    cursor = Write.get_cursor(doc)
                    Write.append_para(cursor=cursor, text=p_txt)

                    Lo.delay(1_000)

                    Lo.close_doc(doc)

                return 0


            def create_para_style(doc: XTextDocument, style_name: str) -> bool:
                try:

                    # font style for the paragraph.
                    font = Font(
                        name="Liberation Serif", size=12.0, color=StandardColor.PURPLE_DARK2, b=True
                    )

                    # spacing below paragraph
                    spc = Spacing(below=UnitMM(4))

                    # paragraph line spacing
                    ln_spc = LineSpacing(mode=ModeKind.FIXED, value=UnitMM(6))

                    # paragraph background color
                    bg_color = ParaBgColor(color=StandardColor.GREEN_LIGHT2)

                    _ = Write.create_style_para(
                        text_doc=doc,
                        style_name=style_name,
                        styles=[font, spc, ln_spc, bg_color]
                    )

                    return True

                except Exception as e:
                    print("Could not set paragraph style")
                    print(f"  {e}")
                return False


            if __name__ == "__main__":
                SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Examples
--------

.. _help_writer_format_style_para_custom_style:

Style a paragraph with a custom style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Create a new paragraph style, assign it to a paragraph and set the paragraph style.

See :ref:`ch06_apply_style_para` for more information.

.. tabs::

    .. code-tab:: python

        # ... other code
        new_style_name = "CoolParagraph"
        if not create_para_style(doc, new_style_name):
            raise RuntimeError(f"Could not create new paragraph style: {new_style_name}")

        # get the XTextRange of the document
        xtext_range = doc.getText().getStart()
        # Load the paragraph style and apply it to the text range.
        para_style = StylePara(new_style_name)
        para_style.apply(xtext_range)
        # ... other code

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229863228-43278981-f7d0-450b-9e0a-cff0f5e43915:
    .. figure:: https://user-images.githubusercontent.com/4193389/229863228-43278981-f7d0-450b-9e0a-cff0f5e43915.png
        :alt: Newly styled paragraph
        :figclass: align-center
        :width: 550px

        Newly styled paragraph.

.. _help_writer_format_style_para_reset_default:

Resetting Paragraph Style to Default
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

By accessing the static property :py:attr:`Para.default <ooodev.format.writer.style.Para.default>` the default paragraph style can be accessed
which is a default style that is set to default values.

This can be used to reset the paragraph style to the default style.


.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.style import Para as StylePara
        # ... other code

        StylePara.default.apply(cursor)
        Write.append_para(cursor=cursor, text="Back to default style.")

        # ... other code

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229866420-8a97695e-1b95-4db0-afe4-a54eaf8eddc4:
    .. figure:: https://user-images.githubusercontent.com/4193389/229866420-8a97695e-1b95-4db0-afe4-a54eaf8eddc4.png
        :alt: Paragraph style reset to default
        :figclass: align-center
        :width: 550px

        Paragraph style reset to default.

Getting style from paragraph
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

By accessing the static property :py:attr:`Para.default <ooodev.format.writer.style.Para.default>` the default paragraph style can be accessed
which is a default style that is set to default values.

This can be used to reset the paragraph style to the default style.


.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.style import Para as StylePara
        # ... other code

        para_cursor = Write.get_paragraph_cursor(cursor)
        para_cursor.gotoPreviousParagraph(False)
        para_cursor.gotoEndOfParagraph(True)
        para_style = StylePara.from_obj(para_cursor)
        # assert name of paragraph style is CoolParagraph
        assert para_style.prop_name == new_style_name

        # ... other code

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`ch06`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - |story_creator|_
        - :py:class:`~ooodev.units.UnitMM`
        - :py:mod:`ooodev.utils.color`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.style.Para`

.. |story_creator| replace:: Story Creator
.. _story_creator: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_story_creator