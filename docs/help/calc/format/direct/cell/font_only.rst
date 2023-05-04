.. _help_calc_format_direct_cell_font_only:

Calc Direct Cell FontOnly
=========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.calc.direct.cell.font.FontOnly` class gives you the same options
as Calc's Font Dialog, but without the dialog. as seen in :numref:`235984034-3859d73c-70c8-4623-9f58-1e5d4a792674`.

.. cssclass:: screen_shot

    .. _235984034-3859d73c-70c8-4623-9f58-1e5d4a792674:
    .. figure:: https://user-images.githubusercontent.com/4193389/235984034-3859d73c-70c8-4623-9f58-1e5d4a792674.png
        :alt: Calc Format Cell dialog Font Effects
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Font Effects


Apply the font to a cell
------------------------

Setup
^^^^^

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 16, 17

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.direct.cell.font import FontOnly

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                cell = Calc.get_cell(sheet=sheet, cell_name="A1")
                font_style = FontOnly(name="Lucida Calligraphy", size=20, font_style="italic")
                Calc.set_val(value="Hello", cell=cell, styles=[font_style])

                f_style = FontOnly.from_obj(cell)
                assert f_style.prop_name == "Lucida Calligraphy"

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the font
^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        font_style = FontOnly(name="Lucida Calligraphy", size=20, font_style="italic")
        Calc.set_val(value="Hello", cell=cell, styles=[font_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236008924-edb77848-d3e9-479a-816b-e6b46296fc6b`.

.. cssclass:: screen_shot

    .. _236008924-edb77848-d3e9-479a-816b-e6b46296fc6b:
    .. figure:: https://user-images.githubusercontent.com/4193389/236008924-edb77848-d3e9-479a-816b-e6b46296fc6b.png
        :alt: Calc Format Cell dialog Font set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Font set


Getting the font from a cell
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = FontOnly.from_obj(cell)
        assert f_style.prop_name == "Lucida Calligraphy"

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply the font to a range
-------------------------

Setup
^^^^^

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 19, 20

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.direct.cell.font import FontOnly

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                Calc.set_val(value="Hello", sheet=sheet, cell_name="A1")
                Calc.set_val(value="World", sheet=sheet, cell_name="B1")
                rng = Calc.get_cell_range(sheet=sheet, range_name="A1:B1")

                font_style = FontOnly(name="Lucida Calligraphy", size=20, font_style="italic")
                font_style.apply(rng)

                f_style = FontOnly.from_obj(rng)
                assert f_style.prop_name == "Lucida Calligraphy"

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the font
^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python
    

        # ... other code
        font_style = FontOnly(name="Lucida Calligraphy", size=20, font_style="italic")
        font_style.apply(rng)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236008924-edb77848-d3e9-479a-816b-e6b46296fc6b`.


Getting the font from a range
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = FontOnly.from_obj(rng)
        assert f_style.prop_name == "Lucida Calligraphy"

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_direct_cell_font`
        - :ref:`help_calc_format_direct_cell_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.direct.cell.font.FontOnly`
        - :py:meth:`Calc.get_cell_range() <ooodev.office.calc.Calc.get_cell_range>`
        - :py:meth:`Calc.get_cell() <ooodev.office.calc.Calc.get_cell>`