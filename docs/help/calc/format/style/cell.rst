.. _help_calc_format_style_cell:

Calc Style Cell / Range
=======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2


Overview
--------

Applying Cell Styles can be accomplished using the ``style_by_name()`` method.

The :py:class:`~ooodev.format.calc.style.cell.kind.style_cell_kind.StyleCellKind` enum is used to lookup the style to be applied.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.format.calc.style import StyleCellKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                cell = sheet["A1"]
                cell.value = "Hello"
                cell.style_by_name(StyleCellKind.ACCENT_1)

                name = cell.style_by_name_get()
                assert name == str(StyleCellKind.ACCENT_1)

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Working with Cell Styles
------------------------

Apply Style to a Cell
^^^^^^^^^^^^^^^^^^^^^

A cell style can be applied while setting the value of a cell.

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.style import StyleCellKind

        # ... other code
        cell = sheet["A1"]
        cell.value = "Hello"
        cell.style_by_name(StyleCellKind.ACCENT_1)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Result be seen in :numref:`236701649-e7c2b254-9e82-4078-abba-9f73a792333d_1`.

.. cssclass:: screen_shot

    .. _236701649-e7c2b254-9e82-4078-abba-9f73a792333d_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236701649-e7c2b254-9e82-4078-abba-9f73a792333d.png
        :alt: Style applied to Cell
        :figclass: align-center
        :width: 550px

        Style applied to Cell

Get Style from a Cell
^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        name = cell.style_by_name_get()
        assert name == str(StyleCellKind.ACCENT_1)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set Default Style for a Cell
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

By calling ``style_by_name()`` without any arguments, the default style will be applied.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell.style_by_name()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Working with Cell Range Styles
------------------------------

Apply Style to a Range
^^^^^^^^^^^^^^^^^^^^^^

In this example we will set the values of a range and apply a style.

.. tabs::

    .. code-tab:: python

        # ... other code
        rng = sheet.rng("A1:B1")
        sheet.set_array(values=[[101, 22]], range_obj=rng)
        cell_rng = sheet.get_range(range_obj=rng)
        cell_rng.style_by_name(StyleCellKind.ACCENT_1)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Result be seen in :numref:`236703868-df15c8d5-08ef-492a-be04-dc7dbbde410e_1`.


.. cssclass:: screen_shot

    .. _236703868-df15c8d5-08ef-492a-be04-dc7dbbde410e_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236703868-df15c8d5-08ef-492a-be04-dc7dbbde410e.png
        :alt: Style applied to Cell
        :figclass: align-center
        :width: 550px

        Style applied to Cell

Get Style from a Range
^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        name = cell_rng.style_by_name_get()
        assert name == str(StyleCellKind.ACCENT_1)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set Default Style for a Cell Range
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

By calling ``style_by_name()`` without any arguments, the default style will be applied.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_rng.style_by_name()

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
        - :ref:`help_calc_format_style_static_cell`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.style.Cell`
