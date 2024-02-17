.. _help_calc_format_style_static_cell:

Calc Style Cell
===============

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2


Overview
--------

Applying Cell/Range Styles can be accomplished using the :py:class:`ooodev.format.calc.style.Cell` class.

The :py:class:`~ooodev.format.calc.style.cell.kind.style_cell_kind.StyleCellKind` enum is used to lookup the style to be applied.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.calc.style import Cell as CellStyle, StyleCellKind
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                sheet = Calc.get_active_sheet()

                style = CellStyle(name=StyleCellKind.ACCENT_1)
                cell_obj = Calc.get_cell_obj("A1")

                Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj, styles=[style])
                cell = Calc.get_cell(sheet=sheet, cell_obj=cell_obj)

                style_obj = CellStyle.from_obj(cell)
                assert style_obj.prop_name == str(StyleCellKind.ACCENT_1)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply Style to a Cell
---------------------

Apply while setting Value
^^^^^^^^^^^^^^^^^^^^^^^^^

A cell style can be applied while setting the value of a cell.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("A1")

        style = CellStyle(name=StyleCellKind.ACCENT_1)
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj, styles=[style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Result be seen in :numref:`236701649-e7c2b254-9e82-4078-abba-9f73a792333d`.

Apply using Calc.set_style_cell()
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`Calc.set_style_cell() <ooodev.office.calc.Calc.set_style_cell>` method can be used to apply one or more styles to a cell.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("A1")

        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        Calc.set_style_cell(sheet=sheet, cell_obj=cell_obj, styles=[style])


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Result be seen in :numref:`236701649-e7c2b254-9e82-4078-abba-9f73a792333d`.


Apply Style to a XCell
^^^^^^^^^^^^^^^^^^^^^^

A cell style can be applied while to an existing cell object
by getting the ``XCell`` object and applying the style.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("A1")

        style = CellStyle(name=StyleCellKind.ACCENT_1)
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet=sheet, cell_obj=cell_obj)
        style.apply(cell)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Result be seen in :numref:`236701649-e7c2b254-9e82-4078-abba-9f73a792333d`.

.. cssclass:: screen_shot

    .. _236701649-e7c2b254-9e82-4078-abba-9f73a792333d:

    .. figure:: https://user-images.githubusercontent.com/4193389/236701649-e7c2b254-9e82-4078-abba-9f73a792333d.png
        :alt: Style applied to Cell
        :figclass: align-center
        :width: 550px

        Style applied to Cell

Get Style from a Cell
^^^^^^^^^^^^^^^^^^^^^

Get Style from a Cell by first getting the ``XCell`` object and then calling ``CellStyle.from_obj()``
passing in the ``XCell`` object.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell = Calc.get_cell(sheet=sheet, cell_obj=cell_obj)

        style_obj = CellStyle.from_obj(cell)
        assert style_obj.prop_name == str(StyleCellKind.ACCENT_1)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply Style to a Range
----------------------

Apply while setting Array
^^^^^^^^^^^^^^^^^^^^^^^^^

A cell style can be applied while setting an array of values.

In this example we will set the values of a range and apply a style.

.. tabs::

    .. code-tab:: python

        # ... other code
        style = CellStyle(name=StyleCellKind.ACCENT_1)
        cell_rng = Calc.get_range_obj(range_name="A1:B1")
        Calc.set_array(values=[[101, 22]], sheet=sheet, range_obj=cell_rng, styles=[style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Result be seen in :numref:`236703868-df15c8d5-08ef-492a-be04-dc7dbbde410e`.


Apply using Calc.set_style_range()
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`Calc.set_style_range() <ooodev.office.calc.Calc.set_style_range>` method can be used to apply one or more styles to a range.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_rng = Calc.get_range_obj(range_name="A1:B1")

        Calc.set_array(values=[[101, 22]], sheet=sheet, range_obj=cell_rng)
        Calc.set_style_range(sheet=sheet, range_obj=cell_rng, styles=[style])


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Result be seen in :numref:`236703868-df15c8d5-08ef-492a-be04-dc7dbbde410e`.


Apply Style to a XCellRange
^^^^^^^^^^^^^^^^^^^^^^^^^^^

In this example we will set the values of a range and apply a style to the range.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_rng = Calc.get_range_obj(range_name="A1:B1")

        Calc.set_array(values=[[101, 22]], sheet=sheet, range_obj=cell_rng)
        rng = Calc.get_cell_range(sheet=sheet, range_obj=cell_rng)
        style.apply(rng)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Result be seen in :numref:`236703868-df15c8d5-08ef-492a-be04-dc7dbbde410e`.

.. cssclass:: screen_shot

    .. _236703868-df15c8d5-08ef-492a-be04-dc7dbbde410e:

    .. figure:: https://user-images.githubusercontent.com/4193389/236703868-df15c8d5-08ef-492a-be04-dc7dbbde410e.png
        :alt: Style applied to Cell
        :figclass: align-center
        :width: 550px

        Style applied to Cell

Get Style from a Range
^^^^^^^^^^^^^^^^^^^^^^

Get Style from a Cell by first getting the ``XCell`` object and then calling ``CellStyle.from_obj()``
passing in the ``XCell`` object.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_rng = Calc.get_range_obj(range_name="A1:B1")
        rng = Calc.get_cell_range(sheet=sheet, range_obj=cell_rng)

        style_obj = CellStyle.from_obj(rng)
        assert style_obj.prop_name == str(StyleCellKind.ACCENT_1)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`ch06`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.style.Cell`
        - :py:class:`Calc.set_val() <ooodev.office.calc.Calc.set_val>`
        - :py:class:`Calc.set_array() <ooodev.office.calc.Calc.set_array>`
        - :py:class:`Calc.set_array_range() <ooodev.office.calc.Calc.set_array_range>`
        - :py:class:`Calc.set_style_cell() <ooodev.office.calc.Calc.set_style_cell>`
        - :py:class:`Calc.set_style_range() <ooodev.office.calc.Calc.set_style_range>`
