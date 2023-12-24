.. _help_writer_format_modify_para_transparency:

Write Modify Paragraph Transparency
===================================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.transparency.Transparency` and :py:class:`ooodev.format.writer.modify.para.transparency.Gradient`
classes are used to modify the transparency values seen in :numref:`234730425-52bcadf8-d827-4198-93a0-8bfb4168e609` of a paragraph style.


Default Paragraph Transparency Style Dialog

.. cssclass:: screen_shot

    .. _234730425-52bcadf8-d827-4198-93a0-8bfb4168e609:

    .. figure:: https://user-images.githubusercontent.com/4193389/234730425-52bcadf8-d827-4198-93a0-8bfb4168e609.png
        :alt: Writer dialog Paragraph Transparency default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Transparency default


Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.para.transparency import Transparency, Gradient, StyleParaKind
        from ooodev.format.writer.modify.para.transparency import GradientStyle, IntensityRange
        from ooodev.format.writer.modify.para.area import Color as StyleAreaColor
        from ooodev.format import Styler
        from ooodev.utils.color import StandardColor
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_kind = StyleParaKind.STANDARD
                para_color_style = StyleAreaColor(color=StandardColor.BLUE_LIGHT2, style_name=para_kind)
                para_transparency_style = Transparency(value=52, style_name=para_kind)
                Styler.apply(doc, para_color_style, para_transparency_style)

                style_obj = Transparency.from_style(doc=doc, style_name=para_kind)
                assert style_obj.prop_style_name == str(para_kind)
                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Transparency
------------

Setting Transparency
^^^^^^^^^^^^^^^^^^^^

Note that we first set a color for the paragraph style. This is because the transparency is not visible unless there is a color.

.. tabs::

    .. code-tab:: python

        # ... other code

        para_kind = StyleParaKind.STANDARD
        para_color_style = StyleAreaColor(color=StandardColor.BLUE_LIGHT2, style_name=para_kind)
        para_transparency_style = Transparency(value=52, style_name=para_kind)
        Styler.apply(doc, para_color_style, para_transparency_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234732332-0c3f5ce4-ee03-4719-b3c1-737c8f9ce081:

    .. figure:: https://user-images.githubusercontent.com/4193389/234732332-0c3f5ce4-ee03-4719-b3c1-737c8f9ce081.png
        :alt: Writer dialog Paragraph Transparency style Transparency changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Transparency style Transparency changed


Getting transparency from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Transparency.from_style(doc=doc, style_name=para_kind)
                assert style_obj.prop_style_name == str(para_kind)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Gradient
--------

Setting Gradient
^^^^^^^^^^^^^^^^

Note that we first set a color for the paragraph style. This is because the gradient is not visible unless there is a color.

.. tabs::

    .. code-tab:: python

        # ... other code

        para_kind = StyleParaKind.STANDARD
        para_color_style = StyleAreaColor(color=StandardColor.BLUE_LIGHT2, style_name=para_kind)
        para_gradient_style = Gradient(
            style=GradientStyle.LINEAR,
            angle=45,
            border=22,
            grad_intensity=IntensityRange(0, 100),
            style_name=para_kind,
        )
        Styler.apply(doc, para_color_style, para_gradient_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234733094-02ec8616-679e-40e0-9e2f-951764b0a0e9:

    .. figure:: https://user-images.githubusercontent.com/4193389/234733094-02ec8616-679e-40e0-9e2f-951764b0a0e9.png
        :alt: Writer dialog Paragraph Transparency style Gradient changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Transparency style Gradient changed

Getting gradient from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Gradient.from_style(doc=doc, style_name=para_kind)
        assert style_obj.prop_style_name == str(para_kind)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_para_transparency`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.transparency.Transparency`
        - :py:class:`ooodev.format.writer.modify.para.transparency.Gradient`