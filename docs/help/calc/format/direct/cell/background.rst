.. _help_calc_format_direct_cell_background:

Calc Direct Cell Background
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Color
-----

The ``style_area_color()`` method is called to set the background color of a cell.

Apply the background color to a cell
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Setup
"""""

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.utils.color import StandardColor

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                cell = sheet["A1"]
                cell.value = "Hello"
                cell.style_area_color(StandardColor.BLUE_LIGHT2)

                Lo.delay(1_000)
                doc.close()
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

        # ... other code
        cell = sheet["A1"]
        cell.value = "Hello"
        cell.style_area_color(StandardColor.BLUE_LIGHT2)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236345774-58cab818-ecde-4211-b4df-d4d239055aea_1` and :numref:`236345626-40553451-919f-4f8f-ae0f-85e19fcf54c8_1`.

.. cssclass:: screen_shot

    .. _236345774-58cab818-ecde-4211-b4df-d4d239055aea_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236345774-58cab818-ecde-4211-b4df-d4d239055aea.png
        :alt: Calc Cell Background Color set
        :figclass: align-center
        :width: 450px

        Calc Cell Background Color set

    .. _236345626-40553451-919f-4f8f-ae0f-85e19fcf54c8_1:

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
        f_style = cell.style_area_color_get()
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

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.utils.color import StandardColor


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                rng = sheet.rng("A1:B1")
                sheet.set_array(
                    values=[["Hello", "World"]], range_obj=rng
                )
                cell_rng = sheet.get_range(range_obj=rng)
                cell_rng.style_area_color(StandardColor.BLUE_LIGHT2)

                Lo.delay(1_000)
                doc.close()
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

        # ... other code
        cell_rng.style_area_color(StandardColor.BLUE_LIGHT2)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236345626-40553451-919f-4f8f-ae0f-85e19fcf54c8` and :numref:`236353030-560861c1-7f6a-4954-b913-81735c139a90_1`.

.. cssclass:: screen_shot

    .. _236353030-560861c1-7f6a-4954-b913-81735c139a90_1:

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

        f_style = cell_rng.style_area_color_get()
        assert f_style.prop_color == StandardColor.BLUE_LIGHT2

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_modify_cell_background`
        - :py:class:`~ooodev.loader.Lo`