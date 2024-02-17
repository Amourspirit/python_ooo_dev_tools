.. _help_calc_format_direct_cell_font_only:

Calc Direct Cell Font Only
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.calc.direct.cell.font.FontOnly` class gives you the same options
as Calc's Font Dialog, but without the dialog. as seen in :numref:`235984034-3859d73c-70c8-4623-9f58-1e5d4a792674_1`.

.. cssclass:: screen_shot

    .. _235984034-3859d73c-70c8-4623-9f58-1e5d4a792674_1:

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

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                cell = sheet["A1"]
                cell.value = "Hello"
                cell.style_font(
                    name="Lucida Calligraphy",
                    size=20,
                    font_style="italic",
                )

                f_style = cell.style_font_get()
                assert f_style is not None

                Lo.delay(1_000)
                doc.close()
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
        cell = sheet["A1"]
        cell.value = "Hello"
        cell.style_font(
            name="Lucida Calligraphy",
            size=20,
            font_style="italic",
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236008924-edb77848-d3e9-479a-816b-e6b46296fc6b_1`.

.. cssclass:: screen_shot

    .. _236008924-edb77848-d3e9-479a-816b-e6b46296fc6b_1:

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

        f_style = cell.style_font_get()
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

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                rng = sheet.rng("A1:B1")
                sheet.set_array(values=[["Hello", "World"]], range_obj=rng)

                cell_rng = sheet.get_range(range_obj=rng)
                cell_rng.style_font(
                    name="Lucida Calligraphy",
                    size=20,
                    font_style="italic",
                )

                Lo.delay(1_000)
                doc.close()
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
        cell_rng = sheet.get_range(range_obj=rng)
        cell_rng.style_font(
            name="Lucida Calligraphy",
            size=20,
            font_style="italic",
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236008924-edb77848-d3e9-479a-816b-e6b46296fc6b`.


Getting the font from a range
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = cell_rng.style_font_get()
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
        - :ref:`help_calc_format_modify_cell_font_only`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.direct.cell.font.FontOnly`