.. _help_writer_format_direct_para_text_flow:

Write Direct Paragraph Text Flow
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Specify hyphenation and pagination options.

Writer has an Text Flow dialog tab.

The :py:class:`ooodev.format.writer.direct.para.text_flow.Breaks`, :py:class:`ooodev.format.writer.direct.para.text_flow.FlowOptions`
and :py:class:`ooodev.format.writer.direct.para.text_flow.Hyphenation` classes is used to set the paragraph text flow.


.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_text_flow:
    .. figure:: https://user-images.githubusercontent.com/4193389/230211906-8ad2ad52-b213-40d1-9752-7ffcf02c2f5e.png
        :alt: Writer Paragraph Text Flow dialog
        :figclass: align-center
        :width: 450px

        Writer Paragraph Text Flow dialog.

Setup
-----

General function used to run these examples:

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from ooodev.office.write import Write
            from ooodev.utils.gui import GUI
            from ooodev.utils.lo import Lo
            from ooodev.format.writer.direct.para.text_flow import Breaks, BreakType, FlowOptions, Hyphenation
            
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
                    tf = Hyphenation(auto=True)
                    Write.append_para(cursor=cursor, text=p_txt, styles=[tf])

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

Generally, if paragraph alignment is set using :py:class:`~.write.Write` then the alignment style only applies to current paragraph.

Auto Hyphenation
----------------

Apply automatic hyphenation.

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        tf = Hyphenation(auto=True)
        Write.append_para(cursor=cursor, text=p_txt, styles=[tf])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230214251-e2d46497-d564-4e2d-828a-a91c4dc8fccc:
    .. figure:: https://user-images.githubusercontent.com/4193389/230214251-e2d46497-d564-4e2d-828a-a91c4dc8fccc.png
        :alt: Writer Style Inspector
        :figclass: align-center

        Writer Style Inspector with the paragraph text flow settings.


Set Flow Options
----------------

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        flow_opt = FlowOptions(orphans=3, widows=4, keep=True)
        Write.append_para(cursor=cursor, text=p_txt, styles=[flow_opt])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230216245-e5636ac2-beac-4057-9f18-cfdd0d0e8c1d:
    .. figure:: https://user-images.githubusercontent.com/4193389/230216245-e5636ac2-beac-4057-9f18-cfdd0d0e8c1d.png
        :alt: Writer Style Inspector
        :figclass: align-center

        Writer Style Inspector with the paragraph text flow settings.


Set Breaks
----------

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text="Set Break in next paragraph...")
        tf_breaks = Breaks(type=BreakType.PAGE_BEFORE, style="Right Page")
        Write.append_para(cursor=cursor, text=p_txt, styles=[tf_breaks])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _230217953-858d05b2-48b7-4c2a-87bd-2ffde9546a0c:
    .. figure:: https://user-images.githubusercontent.com/4193389/230217953-858d05b2-48b7-4c2a-87bd-2ffde9546a0c.png
        :alt: Writer Style Inspector
        :figclass: align-center

        Writer Style Inspector with the paragraph text flow settings.

.. seealso::

   .. cssclass:: ul-list

        - :ref:`help_writer_format_style_para_reset_default`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.text_flow.Breaks`
        - :py:class:`ooodev.format.writer.direct.para.text_flow.FlowOptions`
        - :py:class:`ooodev.format.writer.direct.para.text_flow.Hyphenation`
