.. _help_writer_format_modify_para_pattern:

Write Modify Paragraph Area Pattern
===================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.area.Pattern` class is used to modify the values seen in :numref:`234429874-d86ddec0-381e-4411-854f-6f2f8148de2e` of a paragraph style.

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
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_area_pattern_style = ParaStylePattern.from_preset(
                    preset=PresetPatternKind.SHINGLE, style_name=StyleParaKind.STANDARD
                )
                para_area_pattern_style.apply(doc)

                style_obj = ParaStyleImg.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply pattern to a style
------------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234429874-d86ddec0-381e-4411-854f-6f2f8148de2e:

    .. figure:: https://user-images.githubusercontent.com/4193389/234429874-d86ddec0-381e-4411-854f-6f2f8148de2e.png
        :alt: Writer dialog Paragraph Area Pattern style default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Area Pattern style default

Apply style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_area_pattern_style = ParaStylePattern.from_preset(
            preset=PresetPatternKind.SHINGLE, style_name=StyleParaKind.STANDARD
        )
        para_area_img_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After appling style
^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234428550-31533a46-102b-4a1b-99cf-4cb2d5eb6e19:

    .. figure:: https://user-images.githubusercontent.com/4193389/234428550-31533a46-102b-4a1b-99cf-4cb2d5eb6e19.png
        :alt: Writer dialog Paragraph Area Pattern style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Pattern style changed


Getting the area pattern from a style
-------------------------------------

We can get the area pattern from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = ParaStylePattern.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_direct_para_area_pattern`
        - :ref:`help_writer_format_modify_page_area`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.modify.para.area.Pattern`