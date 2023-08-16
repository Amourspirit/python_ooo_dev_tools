.. _help_writer_format_modify_para_image:

Write Modify Paragraph Area Image
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.area.Img` class is used to modify the values seen in :numref:`234425427-ee1a2151-43ea-4954-ace6-2f872604363a` of a paragraph style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 11, 12, 13, 14

        from ooodev.format.writer.modify.para.area import Gradient as ParaStyleGradient, StyleParaKind
        from ooodev.format.writer.modify.para.area import PresetImageKind

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_area_img_style = ParaStyleImg.from_preset(
                    preset=PresetImageKind.COFFEE_BEANS, style_name=StyleParaKind.STANDARD
                )
                para_area_img_style.apply(doc)

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

Apply image to a style
----------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234425427-ee1a2151-43ea-4954-ace6-2f872604363a:

    .. figure:: https://user-images.githubusercontent.com/4193389/234425427-ee1a2151-43ea-4954-ace6-2f872604363a.png
        :alt: Writer dialog Paragraph Area Image style default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Area Image style default

Apply style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_area_img_style = ParaStyleImg.from_preset(
            preset=PresetImageKind.COFFEE_BEANS, style_name=StyleParaKind.STANDARD
        )
        para_area_img_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After applying style
^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234425641-e6893b4c-9c46-45ba-9852-b155a7a503dd:

    .. figure:: https://user-images.githubusercontent.com/4193389/234425641-e6893b4c-9c46-45ba-9852-b155a7a503dd.png
        :alt: Writer dialog Paragraph Area Image style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Image style changed


Getting the area image from a style
-----------------------------------

We can get the area image from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = ParaStyleImg.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_direct_para_area_gradient`
        - :ref:`help_writer_format_direct_para_area_img`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.area.Img`