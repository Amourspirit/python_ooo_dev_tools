.. _help_writer_format_modify_para_gradient:

Write Modify Paragraph Area Gradient
====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.area.Gradient` class is used to modify the values seen in :numref:`234417354-92e0d839-5fcd-414f-8b5f-487aece63b03` of a paragraph style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 11, 12, 13, 14

        from ooodev.format.writer.modify.para.area import Gradient as ParaStyleGradient, StyleParaKind
        from ooodev.format.writer.modify.para.area import PresetGradientKind

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_area_gradient_style = ParaStyleGradient.from_preset(
                    preset=PresetGradientKind.GREEN_GRASS, style_name=StyleParaKind.STANDARD
                )
                para_area_gradient_style.apply(doc)

                style_obj = ParaStyleGradient.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply alignment to a style
--------------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234417354-92e0d839-5fcd-414f-8b5f-487aece63b03:

    .. figure:: https://user-images.githubusercontent.com/4193389/234417354-92e0d839-5fcd-414f-8b5f-487aece63b03.png
        :alt: Writer dialog Paragraph Area Gradient style default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Area Gradient style default

Apply style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_area_gradient_style = ParaStyleGradient.from_preset(
            preset=PresetGradientKind.GREEN_GRASS, style_name=StyleParaKind.STANDARD
        )
        para_area_gradient_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After appling style
^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234418293-d0282e9a-8183-4015-9d62-eb72cda84a09:

    .. figure:: https://user-images.githubusercontent.com/4193389/234418293-d0282e9a-8183-4015-9d62-eb72cda84a09.png
        :alt: Writer dialog Paragraph Area Gradient style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Gradient style changed


Getting the area gradient from a style
--------------------------------------

We can get the area gradient from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = ParaStyleGradient.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_modify_page_area`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.area.Gradient`