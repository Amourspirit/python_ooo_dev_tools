.. _help_calc_format_direct_cell_font:

Calc Direct Cell Font Class
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.utils.color import CommonColor

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(400)

                a1 = sheet["A1"]
                a1.value = "Hello"
                a1.style_font_general(
                    b=True, color=CommonColor.DARK_GREEN
                )

                b1 = a1.get_cell_right()
                b1.value = "World"
                b1.style_font_general(
                    b=True, u=True, color=CommonColor.DARK_GREEN
                )

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Examples
--------

Set Text Font
+++++++++++++

.. tabs::

    .. code-tab:: python

        a1 = sheet["A1"]
        a1.value = "Hello"
        a1.style_font_general(
            b=True, color=CommonColor.DARK_GREEN
        )

        b1 = a1.get_cell_right()
        b1.value = "World"
        b1.style_font_general(
            b=True, u=True, color=CommonColor.DARK_GREEN
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210860379-03298ed5-1228-42fb-8f04-f7749821a755:
    .. figure:: https://user-images.githubusercontent.com/4193389/210860379-03298ed5-1228-42fb-8f04-f7749821a755.png
        :alt: Styled Text
        :figclass: align-center

        Styled Text

Set Font based upon values
++++++++++++++++++++++++++

.. tabs::

    .. code-tab:: python


        from ooodev.format.calc.direct.cell.font import Font
        import random
        # ... other code

        num_rng = 5
        data = [[random.randint(-100, 100) for _ in range(num_rng)] for _ in range(num_rng)]
        sheet.set_array(values=data, name="A1")
        rng_obj = sheet.find_used_range_obj()
        ft_pos = Font(color=CommonColor.DARK_GREEN, b=True)
        ft_neg = ft_pos.fmt_color(CommonColor.DARK_RED).underline
        with doc:
            # lock controllers for faster processing and avoid flickering.
            for cell_objs in rng_obj.get_cells():
                for co in cell_objs:
                    cell = sheet[co]
                    val = cell.get_num()
                    if val < 0:
                        cell.apply_styles(ft_neg)
                    else:
                        cell.apply_styles(ft_pos)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210923541-b27b63bc-9ffc-4324-b88b-9d05dd1e0dc7_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210923541-b27b63bc-9ffc-4324-b88b-9d05dd1e0dc7.png
        :alt: Styled Array
        :figclass: align-center

        Styled array

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_direct_cell_font_only`
        - :ref:`help_calc_format_direct_cell_font_effects`
        - :py:class:`ooodev.format.calc.direct.cell.font.Font`
