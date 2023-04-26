.. _help_writer_format_direct_para_alignment:

Write Direct Paragraph Alignment
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Writer has an Alignment dialog tab.

The :py:class:`ooodev.format.writer.direct.para.alignment.Alignment` class is used to set the paragraph alignment.


.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_alignment:
    .. figure:: https://user-images.githubusercontent.com/4193389/230089934-1d5d6108-cf79-4ae0-b0b6-2aad00f70c57.png
        :alt: Writer Paragraph Alignment dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Alignment dialog.

Setup
-----

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from ooodev.office.write import Write
            from ooodev.utils.gui import GUI
            from ooodev.utils.lo import Lo
            from ooodev.format.writer.direct.para.alignment import Alignment, LastLineKind, ParagraphAdjust
            
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
                    al = Alignment().align_left
                    Write.append_para(cursor=cursor, text=p_txt, styles=[al])
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

Generally, if paragraph alignment is set using :py:class:`~ooodev.office.write.Write` then the alignment style only applies to current paragraph.

Align Default
^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        al_default = Alignment().default
        Write.append_para(cursor=cursor, text=p_txt, styles=[al_default])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230091775-7033e23e-d02c-4ec7-9267-60dae0485ca1:
    .. figure:: https://user-images.githubusercontent.com/4193389/230091775-7033e23e-d02c-4ec7-9267-60dae0485ca1.png
        :alt: Writer Paragraph Alignment Default.
        :figclass: align-center

        Writer Paragraph Alignment Default.


Align Left
^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        al = Alignment().align_left
        Write.append_para(cursor=cursor, text=p_txt, styles=[al])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230092708-a37afe8f-2885-44fe-aeaa-fdcaed16bdd3:
    .. figure:: https://user-images.githubusercontent.com/4193389/230092708-a37afe8f-2885-44fe-aeaa-fdcaed16bdd3.png
        :alt: Writer Paragraph Alignment Left.
        :figclass: align-center

        Writer Paragraph Alignment Left.

Align Center
^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        al = Alignment().align_center
        Write.append_para(cursor=cursor, text=p_txt, styles=[al])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230093384-2d60ae7e-86c8-4edc-ba61-2e55abe4d495:
    .. figure:: https://user-images.githubusercontent.com/4193389/230093384-2d60ae7e-86c8-4edc-ba61-2e55abe4d495.png
        :alt: Writer Paragraph Alignment Center
        :figclass: align-center

        Writer Paragraph Alignment Center.


Align Right
^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        al = Alignment().align_right
        Write.append_para(cursor=cursor, text=p_txt, styles=[al])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230094045-eca4b96d-0a85-4387-ba68-20610030f4c3:
    .. figure:: https://user-images.githubusercontent.com/4193389/230094045-eca4b96d-0a85-4387-ba68-20610030f4c3.png
        :alt: Writer Paragraph Alignment Right
        :figclass: align-center

        Writer Paragraph Alignment Right.


Align Justified
^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        al = Alignment().justified
        Write.append_para(cursor=cursor, text=p_txt, styles=[al])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230094826-4379d34d-ba46-4312-badb-e881bc828350:
    .. figure:: https://user-images.githubusercontent.com/4193389/230094826-4379d34d-ba46-4312-badb-e881bc828350.png
        :alt: Writer Paragraph Alignment Justified
        :figclass: align-center

        Writer Paragraph Alignment Justified.

Justified Last Line
"""""""""""""""""""

It is possible to also set the last line when using Align Justified.

**Justify last line center.**

.. tabs::

    .. code-tab:: python

        # ... other code
        al = Alignment(align_last=LastLineKind.CENTER).justified
        Write.append_para(cursor=cursor, text=p_txt, styles=[al])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230095705-b39e0523-8fb2-4988-a256-fd3ef240eca3:
    .. figure:: https://user-images.githubusercontent.com/4193389/230095705-b39e0523-8fb2-4988-a256-fd3ef240eca3.png
        :alt: Writer Paragraph Alignment Justified, Center
        :figclass: align-center

        Writer Paragraph Alignment Justified, Center.


**Justify last line justified.**

.. tabs::

    .. code-tab:: python

        # ... other code
        al = Alignment(align_last=LastLineKind.JUSTIFY).justified
        Write.append_para(cursor=cursor, text=p_txt, styles=[al])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230096315-bf6a26ee-a846-4151-bf97-5c1bf1f04f4d:
    .. figure:: https://user-images.githubusercontent.com/4193389/230096315-bf6a26ee-a846-4151-bf97-5c1bf1f04f4d.png
        :alt: Writer Paragraph Alignment Justified, JUSTIFY
        :figclass: align-center

        Writer Paragraph Alignment Justified, JUSTIFY.

Set Cursor
----------

Sometimes it may be necessary to set the cursor style for many paragraphs.
The best way to accomplish this is to set the cursor style before appending paragraphs.

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        al = Alignment(align_last=LastLineKind.CENTER).justified
        al.apply(cursor)
        Write.append_para(cursor=cursor, text=p_txt)
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230097129-a5538cb1-6bec-4f6d-8204-a00e19ebfe70:
    .. figure:: https://user-images.githubusercontent.com/4193389/230098129-a5538cb1-6bec-4f6d-8204-a00e19ebfe70.png
        :alt: Writer Paragraph Alignment via cursor
        :figclass: align-center

        Writer Paragraph Alignment via cursor.

For resetting cursor to default, see :ref:`help_writer_format_style_para_reset_default`.

Getting Alignment from Paragraph
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

A paragraph cursor object is used to select the first paragraph in the document.
The paragraph cursor is then used to get the style.

.. tabs::

    .. code-tab:: python

        # ... other code

        cursor = Write.get_cursor(doc)
        al = Alignment().align_right
        Write.append_para(cursor=cursor, text=p_txt, styles=[al])

        para_cursor = Write.get_paragraph_cursor(cursor)
        para_cursor.gotoPreviousParagraph(False)
        para_cursor.gotoEndOfParagraph(True)

        para_align = Alignment.from_obj(para_cursor)
        assert para_align.prop_align == ParagraphAdjust.RIGHT

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

   .. cssclass:: ul-list

        - :ref:`help_writer_format_style_para_reset_default`
        - :ref:`help_writer_format_style_para`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_para_alignment`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.alignment.Alignment`