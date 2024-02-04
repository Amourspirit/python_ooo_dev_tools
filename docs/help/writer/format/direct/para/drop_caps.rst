.. _help_writer_format_direct_para_drop_caps:

Write Direct Paragraph DropCaps Class
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.writer.direct.para.drop_caps.DropCaps` class can apply the same formatting as Writer's Drop Caps Dialog, :numref:`ss_writer_dialog_para_drop_caps`.

Note: ``DropCaps`` class uses Dispatch Commands. This means the ``DropCaps`` class is not suitable in ``Headless`` mode.

.. cssclass:: screen_shot invert

    .. _ss_writer_dialog_para_drop_caps:
    .. figure:: https://user-images.githubusercontent.com/4193389/212792304-f523ef05-4e77-4743-9dc3-cf2bdc6985f7.png
        :alt: Drop Caps dialog screenshot
        :figclass: align-center
        :width: 450px

        Drop Caps dialog screenshot


Setting the style
-----------------

General function used to run these examples.


.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst


.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from ooodev.office.write import Write
            from ooodev.utils.color import CommonColor
            from ooodev.utils.gui import GUI
            from ooodev.loader.lo import Lo
            from ooodev.format.writer.direct.para.drop_caps import DropCaps, StyleCharKind
            
            def main() -> int:
                p_txt = (
                    |short_ptext|
                )

            with Lo.Loader(Lo.ConnectSocket()):
                doc = Write.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                cursor = Write.get_cursor(doc)
                dc = DropCaps(count=1)
                Write.append_para(cursor=cursor, text=p_txt, styles=[dc])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Examples
--------

Simple Drop Caps
^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        dc = DropCaps(count=1)
        Write.append_para(cursor=cursor, text=p_txt, styles=[dc])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot invert

    .. _212793358-98702597-3f8f-4b72-922f-f654bccdcf47:
    .. figure:: https://user-images.githubusercontent.com/4193389/212793358-98702597-3f8f-4b72-922f-f654bccdcf47.png
        :alt: Drop Caps screenshot
        :figclass: align-center

        Drop Caps screenshot

.. cssclass:: screen_shot invert

    .. _212793414-be74ba43-3ff6-4501-8e73-0ceb304cf2f4:
    .. figure:: https://user-images.githubusercontent.com/4193389/212793414-be74ba43-3ff6-4501-8e73-0ceb304cf2f4.png
        :alt: Drop Caps dialog screenshot
        :figclass: align-center

        Drop Caps dialog screenshot

Apply to Whole Word
^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        dc = DropCaps(count=1, whole_word=True)
        Write.append_para(cursor=cursor, text=p_txt, styles=[dc])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot invert

    .. _212793635-b3e5a4b1-d11f-4bb8-9627-69334ed55f33:
    .. figure:: https://user-images.githubusercontent.com/4193389/212793635-b3e5a4b1-d11f-4bb8-9627-69334ed55f33.png
        :alt: Drop Caps screenshot
        :figclass: align-center

        Drop Caps screenshot

.. cssclass:: screen_shot invert

    .. _212793740-4344d017-4d83-4503-b006-06de1788697a:
    .. figure:: https://user-images.githubusercontent.com/4193389/212793740-4344d017-4d83-4503-b006-06de1788697a.png
        :alt: Drop Caps dialog screenshot
        :figclass: align-center

        Drop Caps dialog screenshot

Increase Drop Caps Spacing
^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        dc = DropCaps(count=1, style=StyleCharKind.DROP_CAPS, spaces=5.0)
        Write.append_para(cursor=cursor, text=p_txt, styles=[dc])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot invert

    .. _212794805-d890a9f5-4c7b-4095-8bdb-f8d85743f8e5:
    .. figure:: https://user-images.githubusercontent.com/4193389/212794805-d890a9f5-4c7b-4095-8bdb-f8d85743f8e5.png
        :alt: Drop Caps screenshot
        :figclass: align-center

        Drop Caps screenshot

.. cssclass:: screen_shot invert

    .. _212794231-348614ea-b3a0-4da5-bf8a-8b3d3327cafa:
    .. figure:: https://user-images.githubusercontent.com/4193389/212794231-348614ea-b3a0-4da5-bf8a-8b3d3327cafa.png
        :alt: Drop Caps dialog screenshot
        :figclass: align-center

        Drop Caps dialog screenshot

Get the Drop Caps Style from a Paragraph
-----------------------------------------

Continuing from the code example above, we can get the Drop Caps from the document.

A paragraph cursor object is used to select the first paragraph in the document.
The paragraph cursor is then used to get the style.

.. tabs::

    .. code-tab:: python

        # ... other code
        para_cursor = Write.get_paragraph_cursor(cursor)
        para_cursor.gotoPreviousParagraph(False)
        para_cursor.gotoEndOfParagraph(True)

        dc = DropCaps.from_obj(para_cursor)
        assert dc.prop_inner.prop_count == 1

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

   .. cssclass:: ul-list

        - :ref:`help_writer_format_style_para_reset_default`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`ch02`
        - :ref:`help_writer_format_modify_para_drop_caps`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.drop_caps.DropCaps`
