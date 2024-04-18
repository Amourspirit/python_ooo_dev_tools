.. _help_calc_format_direct_static_cell_background:

Calc Direct Cell Background (Static)
====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Color
-----

The :py:class:`ooodev.format.calc.direct.cell.background.Color` class is used to set the background color of a cell.

.. seealso::

    - :ref:`help_calc_format_direct_cell_background`

Apply the background color to a cell
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Setup
"""""

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 17, 18

        import uno
        from ooodev.office.calc import Calc
        from ooodev.gui.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.direct.cell.background import Color as BgColor
        from ooodev.utils.color import StandardColor

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                cell = Calc.get_cell(sheet=sheet, cell_name="A1")
                style = BgColor(StandardColor.BLUE_LIGHT2)
                Calc.set_val(value="Hello", cell=cell, styles=[style])

                f_style = BgColor.from_obj(cell)
                assert f_style.prop_color == StandardColor.BLUE_LIGHT2

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the color
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        style = BgColor(StandardColor.BLUE_LIGHT2)
        Calc.set_val(value="Hello", cell=cell, styles=[style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236345774-58cab818-ecde-4211-b4df-d4d239055aea` and :numref:`236345626-40553451-919f-4f8f-ae0f-85e19fcf54c8`.

.. cssclass:: screen_shot

    .. _236345774-58cab818-ecde-4211-b4df-d4d239055aea:

    .. figure:: https://user-images.githubusercontent.com/4193389/236345774-58cab818-ecde-4211-b4df-d4d239055aea.png
        :alt: Calc Cell Background Color set
        :figclass: align-center
        :width: 450px

        Calc Cell Background Color set

    .. _236345626-40553451-919f-4f8f-ae0f-85e19fcf54c8:

    .. figure:: https://user-images.githubusercontent.com/4193389/236345626-40553451-919f-4f8f-ae0f-85e19fcf54c8.png
        :alt: Calc Format Cell dialog Background Color set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Background Color set


Getting the color from a cell
"""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = BgColor.from_obj(cell)
        assert f_style.prop_color == StandardColor.BLUE_LIGHT2

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply the background color to a range
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Setup
"""""

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 17, 18

        import uno
        from ooodev.office.calc import Calc
        from ooodev.gui.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.direct.cell.background import Color as BgColor
        from ooodev.utils.color import StandardColor


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                Calc.set_val(value="Hello", sheet=sheet, cell_name="A1")
                Calc.set_val(value="World", sheet=sheet, cell_name="B1")
                rng = Calc.get_cell_range(sheet=sheet, range_name="A1:B1")

                style = BgColor(StandardColor.BLUE_LIGHT2)
                style.apply(rng)

                f_style = BgColor.from_obj(rng)
                assert f_style.prop_color == StandardColor.BLUE_LIGHT2

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the color
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        style = BgColor(StandardColor.BLUE_LIGHT2)
        style.apply(rng)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236345626-40553451-919f-4f8f-ae0f-85e19fcf54c8` and :numref:`236353030-560861c1-7f6a-4954-b913-81735c139a90`.

.. cssclass:: screen_shot

    .. _236353030-560861c1-7f6a-4954-b913-81735c139a90:

    .. figure:: https://user-images.githubusercontent.com/4193389/236353030-560861c1-7f6a-4954-b913-81735c139a90.png
        :alt: Calc Cell Background Color set
        :figclass: align-center
        :width: 450px

        Calc Cell Background Color set


Getting the color from a range
""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = BgColor.from_obj(rng)
        assert f_style.prop_color == StandardColor.BLUE_LIGHT2

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_calc_format_direct_cell_background`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_modify_cell_background`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`Calc.get_cell_range() <ooodev.office.calc.Calc.get_cell_range>`
        - :py:meth:`Calc.get_cell() <ooodev.office.calc.Calc.get_cell>`