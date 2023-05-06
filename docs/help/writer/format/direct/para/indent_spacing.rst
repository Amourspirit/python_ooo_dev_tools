.. _help_writer_format_direct_para_indent_spacing:

Write Direct Paragraph Indent & Spacing
=======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Writer has a Indents & spacing dialog tab.

The :py:class:`ooodev.format.writer.direct.para.indent_space.Indent`, :py:class:`ooodev.format.writer.direct.para.indent_space.Spacing`,
and :py:class:`ooodev.format.writer.direct.para.indent_space.LineSpacing` classes are used to set the indent, spacing, and line spacing of a paragraph.

.. cssclass:: screen_shot

    .. _ss_writer_dialog_para_indents_and_spacing:
    .. figure:: https://user-images.githubusercontent.com/4193389/229901830-f627ef53-45fd-4533-9aa4-7eecd0762313.png
        :alt: Writer Paragraph Indents and Spacing dialog
        :figclass: align-center
        :width: 450px

        :Writer Paragraph Indents and Spacing dialog

Setup
-----

General function used to run these examples.

.. must be before the tabs directive
.. include:: ../../../../../resources/help/inc/inc_style_short_ptext.rst

.. tabs::

    .. group-tab:: Python

        .. code-block:: python
            :substitutions:

            from ooodev.office.write import Write
            from ooodev.utils.gui import GUI
            from ooodev.utils.lo import Lo
            from ooodev.format.writer.direct.para.indent_space import Indent, Spacing, ModeKind, LineSpacing
            
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
                    indent = Indent(before=22.0, after=20.0, first=8.0)
                    Write.append_para(cursor=cursor, text=p_txt, styles=[indent])
                    Write.append_para(cursor=cursor, text=p_txt)
                    Lo.delay(1_000)
                    Lo.close_doc(doc)
                return 0


            if __name__ == "__main__":
                sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Indent Class
------------

The :py:class:`ooodev.format.writer.direct.para.indent_space.Indent` class is used to set the indent of a paragraph.

Example
^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        indent = Indent(before=22.0, after=20.0, first=8.0)
        Write.append_para(cursor=cursor, text=p_txt, styles=[indent])
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _229902469-875ede25-1466-4787-a4ec-9e3fb0e415ec:
    .. figure:: https://user-images.githubusercontent.com/4193389/229902469-875ede25-1466-4787-a4ec-9e3fb0e415ec.png
        :alt: Indent, before 22.0 mm, after 20.0 mm, first 8.0 mm
        :figclass: align-center

        Indent, before 22.0 mm, after 20.0 mm, first 8.0 mm


Spacing Class
-------------

The :py:class:`ooodev.format.writer.direct.para.indent_space.Spacing` class is used to set the spacing of a paragraph.


Example
^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        spacing = Spacing(above=8.0, below=10.0)
        Write.append_para(cursor=cursor, text=p_txt, styles=[spacing])
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _229905293-8e885095-ae0f-4b7e-9e65-f93101270a62:
    .. figure:: https://user-images.githubusercontent.com/4193389/229905293-8e885095-ae0f-4b7e-9e65-f93101270a62.png
        :alt: Spacing, above 8.0 mm, below 10.0 mm
        :figclass: align-center

        Spacing, above 8.0 mm, below 10.0 mm.

LineSpacing Class
-----------------

The :py:class:`ooodev.format.writer.direct.para.indent_space.LineSpacing` class is used to set the spacing of a paragraph.


Example
^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        ln_spc = LineSpacing(mode=ModeKind.PROPORTIONAL, value=85)
        Write.append_para(cursor=cursor, text=p_txt, styles=[ln_spc])
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _229907266-a4d010bb-a705-41e7-a6ed-72cc2947e108:
    .. figure:: https://user-images.githubusercontent.com/4193389/229907266-a4d010bb-a705-41e7-a6ed-72cc2947e108.png
        :alt: LineSpacing, mode=ModeKind.PROPORTIONAL, value=85
        :figclass: align-center

        LineSpacing, mode=ModeKind.PROPORTIONAL, value=85.

Other Examples
--------------

Apply to start of document
^^^^^^^^^^^^^^^^^^^^^^^^^^

See also: :ref:`ch06_apply_style_para`.

.. tabs::

    .. code-tab:: python

        # ... other code
        # get the XTextRange of the document
        xtext_range = doc.getText().getStart()
        # Created the spacing values apply them to the text range.
        ln_spc = LineSpacing(mode=ModeKind.PROPORTIONAL, value=85)
        indent = Indent(before=8.0, after=8.0, first=8.0)
        Styler.apply(xtext_range, ln_spc, indent)

        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=p_txt)
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _229910058-1ade0357-d9a2-472b-8b67-475ee613baaf:
    .. figure:: https://user-images.githubusercontent.com/4193389/229910058-1ade0357-d9a2-472b-8b67-475ee613baaf.png
        :alt: Direct styles applied to start of document
        :figclass: align-center

        Direct styles applied to start of document.



.. cssclass:: screen_shot

    .. _229910351-a7a029b4-0cb5-40d5-aefe-e38471bb119a:
    .. figure:: https://user-images.githubusercontent.com/4193389/229910351-a7a029b4-0cb5-40d5-aefe-e38471bb119a.png
        :alt: Paragraph Indent & Spacing dialog
        :figclass: align-center
        :width: 450px

        Paragraph Indent & Spacing dialog.


Resetting paragraph styles
^^^^^^^^^^^^^^^^^^^^^^^^^^

Restting paragraph styles is done by applying the :py:attr:`Para.default <ooodev.format.writer.style.Para.default>` styles to the cursor
as seen in :ref:`help_writer_format_style_para_reset_default`.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.style import Para as StylePara
        # ... other code

        # get the XTextRange of the document
        xtext_range = doc.getText().getStart()
        # Created the spacing values apply them to the text range.
        ln_spc = LineSpacing(mode=ModeKind.PROPORTIONAL, value=85)
        indent = Indent(before=8.0, after=8.0, first=8.0)
        Styler.apply(xtext_range, ln_spc, indent)

        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=p_txt)
        # reset the paragraph styles
        StylePara.default.apply(cursor)
        Write.append_para(cursor=cursor, text=p_txt)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _229912041-c1f5afd3-0b01-49bc-a353-265bc4d1e2c6:
    .. figure:: https://user-images.githubusercontent.com/4193389/229912041-c1f5afd3-0b01-49bc-a353-265bc4d1e2c6.png
        :alt: Paragraph styles reset
        :figclass: align-center

        Paragraph styles reset.

Related Topics
--------------

.. seealso::

   .. cssclass:: ul-list

        - :ref:`help_writer_format_style_para_reset_default`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_para_indent_spacing`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.indent_space.Indent`
        - :py:class:`ooodev.format.writer.direct.para.indent_space.Spacing`
        - :py:class:`ooodev.format.writer.direct.para.indent_space.LineSpacing`