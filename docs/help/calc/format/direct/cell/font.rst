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

        from ooodev.utils.lo import Lo
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.format.calc.direct.cell.font import Font
        from ooodev.format import CommonColor, Styler


        def main() -> int:
            with Lo.Loader(Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(300)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_200_PERCENT)

                cell_obj = Calc.get_cell_obj("A1")
                Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
                Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj.right)

                a1 = Calc.get_cell(sheet=sheet, cell_obj=cell_obj)
                b12 = Calc.get_cell(sheet=sheet, cell_obj=cell_obj.right)

                ft = Font(color=CommonColor.DARK_GREEN)
                Styler.apply(a1, ft)
                Styler.apply(b12, ft.bold.underline)
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

Set Text Font
+++++++++++++

.. tabs::

    .. code-tab:: python

        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj.right)

        a1 = Calc.get_cell(sheet=sheet, cell_obj=cell_obj)
        b12 = Calc.get_cell(sheet=sheet, cell_obj=cell_obj.right)

        ft = Font(color=CommonColor.DARK_GREEN)
        Styler.apply(a1, ft)
        Styler.apply(b12, ft.bold.underline)

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

        import random
        # ... other code

        num_rng = 5
        data = [[random.randint(-100, 100) for _ in range(num_rng)] for _ in range(num_rng)]

        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_array(values=data, sheet=sheet, cell_obj=cell_obj)
        rng_obj = Calc.find_used_range_obj(sheet)

        ft_pos = Font(color=CommonColor.DARK_GREEN, b=True)
        ft_neg = ft_pos.fmt_color(CommonColor.DARK_RED).underline

        for cell_objs in rng_obj.get_cells():
            for co in cell_objs:
                val = Calc.get_num(sheet=sheet, cell_obj=co)
                cell = Calc.get_cell(sheet=sheet, cell_obj=co)
                if val < 0:
                    Styler.apply(cell, ft_neg)
                else:
                    Styler.apply(cell, ft_pos)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210923541-b27b63bc-9ffc-4324-b88b-9d05dd1e0dc7:
    .. figure:: https://user-images.githubusercontent.com/4193389/210923541-b27b63bc-9ffc-4324-b88b-9d05dd1e0dc7.png
        :alt: Styled Array
        :figclass: align-center

        Styled array

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_char_font`
        - :py:class:`ooodev.format.calc.direct.cell.font.Font`
