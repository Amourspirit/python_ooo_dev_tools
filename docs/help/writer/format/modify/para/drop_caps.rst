.. _help_writer_format_modify_para_drop_caps:

Write Modify Paragraph Drops Caps
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.drop_caps.DropCaps` class is used to modify the values seen in :numref:`234729276-17e99b85-3e71-44d4-896f-d75563705088` of a paragraph style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 13, 14, 15, 16

        from ooodev.format.writer.modify.para.drop_caps import DropCaps, StyleParaKind, StyleCharKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_dc_style = DropCaps(
                    count=4, style=StyleCharKind.DROP_CAPS, spaces=5.0, style_name=StyleParaKind.STANDARD
                )
                para_dc_style.apply(doc)

                style_obj = DropCaps.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply Drops Caps to a style
---------------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234729276-17e99b85-3e71-44d4-896f-d75563705088:
    .. figure:: https://user-images.githubusercontent.com/4193389/234729276-17e99b85-3e71-44d4-896f-d75563705088.png
        :alt: Writer dialog Paragraph Drop Caps style default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Drop Caps style default

Apply style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        para_dc_style = DropCaps(
            count=4, style=StyleCharKind.DROP_CAPS, spaces=5.0, style_name=StyleParaKind.STANDARD
        )
        para_dc_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After appling style
^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _234729668-176ce8e4-a2b6-475f-9fb5-cad05d713e11:
    .. figure:: https://user-images.githubusercontent.com/4193389/234729668-176ce8e4-a2b6-475f-9fb5-cad05d713e11.png
        :alt: Writer dialog Paragraph Drops Caps style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Drops Caps style changed


Getting the area color from a style
-----------------------------------

We can get the area color from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = DropCaps.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_direct_para_drop_caps`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.drop_caps.DropCaps`