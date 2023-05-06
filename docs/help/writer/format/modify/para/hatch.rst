.. _help_writer_format_modify_para_hatch:

Write Modify Paragraph Area Pattern
===================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.area.Hatch` class is used to modify the values seen in :numref:`234430722-50c899fa-9e11-494b-bd56-daa7989c435a` of a paragraph style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 11, 12, 13, 14

        from ooodev.format.writer.modify.para.area import Pattern as ParaStylePattern
        from ooodev.format.writer.modify.para.area import StyleParaKind, PresetPatternKind

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_area_hatch_style = ParaStyleHatch.from_preset(
                    preset=PresetHatchKind.GREEN_90_DEGREES_TRIPLE, style_name=StyleParaKind.STANDARD
                )
                para_area_hatch_style.apply(doc)

                style_obj = ParaStyleHatch.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply image to a style
----------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234430722-50c899fa-9e11-494b-bd56-daa7989c435a:
    .. figure:: https://user-images.githubusercontent.com/4193389/234430722-50c899fa-9e11-494b-bd56-daa7989c435a.png
        :alt: Writer dialog Paragraph Area Hatch style default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Area Hatch style default

Apply style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_area_hatch_style = ParaStyleHatch.from_preset(
            preset=PresetHatchKind.GREEN_90_DEGREES_TRIPLE, style_name=StyleParaKind.STANDARD
        )
        para_area_hatch_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After appling style
^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234431194-448bedd6-0e3e-44af-88b7-1fc68902f230:
    .. figure:: https://user-images.githubusercontent.com/4193389/234431194-448bedd6-0e3e-44af-88b7-1fc68902f230.png
        :alt: Writer dialog Paragraph Area Hatch style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Hatch style changed


Getting the area hatch from a style
-----------------------------------

We can get the area hatch from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = ParaStyleHatch.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_para_area_hatch`
        - :ref:`help_writer_format_modify_page_area`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.area.Hatch`