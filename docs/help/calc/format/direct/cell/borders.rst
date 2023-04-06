.. _help_calc_format_direct_cell_borders:

Calc Direct Cell Borders
========================

Overview
--------

Specify Calc Cell/Range borders.

Writer has an Text Flow dialog tab.

The :py:class:`ooodev.format.calc.direct.cell.borders.Borders`, :py:class:`ooodev.format.calc.direct.cell.borders.Padding`
and :py:class:`ooodev.format.calc.direct.cell.borders.Shadow` classes is used to set the border values.


.. cssclass:: screen_shot

    .. _ss_calc_format_cell_borders_dialog:
    .. figure:: https://user-images.githubusercontent.com/4193389/230220907-c5c5014c-e701-468b-b811-e7918ff329f6.png
        :alt: Calc Format Cells Borders dialog
        :figclass: align-center
        :width: 450px

        Calc Format Cells Borders dialog.

Setup
-----

General function used to run these examples:

.. tabs::

    .. code-tab:: python

        from ooodev.format import Styler
        from ooodev.format.calc.direct.cell.borders import Borders, Shadow, Side, BorderLineKind, Padding
        from ooodev.office.calc import Calc
        from ooodev.utils.color import CommonColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo


        def main() -> int:
            with Lo.Loader(Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(300)
                Calc.zoom_value(doc, 130)

                rng_obj = Calc.get_range_obj("B2:F6")
                cr = Calc.get_cell_range(sheet, rng_obj)
                borders = Borders(border_side=Side(color=CommonColor.BLUE))
                Styler.apply(cr, borders)
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

Single Cell
^^^^^^^^^^^

Default Border
""""""""""""""

Appling ``Border.default`` will create a default border for a cell or a range.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        Styler.apply(cell, Borders().default)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210101040-aa66cae1-323b-4fb0-b9c2-ac3a82a62403:
    .. figure:: https://user-images.githubusercontent.com/4193389/210101040-aa66cae1-323b-4fb0-b9c2-ac3a82a62403.png
        :alt: Cell with default border
        :figclass: align-center

        Cell with default border.

Removing Border
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        Styler.apply(cell, Borders().default)
        # ...
        # remove border
        Styler.apply(cell, Borders().empty)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Colored border
""""""""""""""
.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        borders = Borders(border_side=Side(color=CommonColor.RED))
        Styler.apply(cell, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210101175-74a38aa2-c77a-4f6c-ad76-3b3f2322c6d9:
    .. figure:: https://user-images.githubusercontent.com/4193389/210101175-74a38aa2-c77a-4f6c-ad76-3b3f2322c6d9.png
        :alt: Cell with colored border
        :figclass: align-center

        Cell with colored border.

Applying border to a side
"""""""""""""""""""""""""

Apply green border to left side.

:py:class:`~ooodev.format.calc.direct.cell.borders.Borders` constructor can also take ``left``, ``right``, ``top``, ``bottom``, ``vertical``, ``horizontal``, ``diagonal_down`` and ``diagonal_up`` arguments as sides.
In this case just pass in the ``left`` side.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        Styler.apply(cell, Borders(left=Side(color=CommonColor.GREEN)))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210101363-4288e162-2117-4b95-bed0-578a179b31f1:
    .. figure:: https://user-images.githubusercontent.com/4193389/210101363-4288e162-2117-4b95-bed0-578a179b31f1.png
        :alt: Cell with left colored border
        :figclass: align-center

        Cell with left colored border.

Apply border with increased size
""""""""""""""""""""""""""""""""

Passing ``width`` argument to ``Side()`` controls border width.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        side_left_right = Side(color=CommonColor.GREEN, width=1.8)
        borders = Borders(left=side_left_right, right=side_left_right)
        Styler.apply(cell, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210101564-b76cd842-ed82-4fd9-85b6-16890cb80364:
    .. figure:: https://user-images.githubusercontent.com/4193389/210100564-b76cd842-ed82-4fd9-85b6-16890cb80364.png
        :alt: Cell with left and right colored border
        :figclass: align-center

        Cell with left and right colored border.

Apply different top and bottom colors
"""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        side_top_bottom = Side(color=CommonColor.CHARTREUSE, width=1.3)
        side_left_right = Side(color=CommonColor.ROYAL_BLUE, width=1.3)
        borders = Borders(
            top=side_top_bottom,
            bottom=side_top_bottom,
            left=side_left_right,
            right=side_left_right,
            )
        Styler.apply(cell, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210102075-e8d7229b-b480-45cf-b8d2-5782d36ac6c8:
    .. figure:: https://user-images.githubusercontent.com/4193389/210102075-e8d7229b-b480-45cf-b8d2-5782d36ac6c8.png
        :alt: Cell with left and right colored border
        :figclass: align-center

        Cell with left and right colored border.

Apply Diagonal border
"""""""""""""""""""""

Passing ``diagonal_up`` and ``diagonal_down`` arguments to :py:class:`~ooodev.format.calc.direct.cell.borders.Borders` allows for diagonal lines.

**UP**

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        borders = Borders(diagonal_up=Side(color=CommonColor.RED))
        Styler.apply(cell, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210102706-ebe79c91-5e0a-4482-a58d-a797efa7ded9:
    .. figure:: https://user-images.githubusercontent.com/4193389/210102706-ebe79c91-5e0a-4482-a58d-a797efa7ded9.png
        :alt: Cell with diagonal up colored border
        :figclass: align-center

        Cell with diagonal up colored border.


**DOWN**

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        borders = Borders(diagonal_down=Side(color=CommonColor.RED))
        Styler.apply(cell, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210102945-73b453d6-33f2-4582-a276-61fda1e5edbe:
    .. figure:: https://user-images.githubusercontent.com/4193389/210102945-73b453d6-33f2-4582-a276-61fda1e5edbe.png
        :alt: Cell with diagonal down colored border
        :figclass: align-center

        Cell with diagonal down colored border.

Applying different style border
"""""""""""""""""""""""""""""""

Using py:class:`~ooodev.format.inner.direct.structs.side.BorderLineKind` enumeration it is possible to change the border style to many different configurations.

In this example the border style is set to Dash-dot.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        borders = Borders(
            border_side=Side(line=BorderLineKind.DASH_DOT, color=CommonColor.DARK_GREEN)
        )
        Styler.apply(cell, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210103415-147a46c0-7d99-4cd4-b861-d46228a89c25:
    .. figure:: https://user-images.githubusercontent.com/4193389/210104415-147a46c0-7d99-4cd4-b861-d46228a89c25.png
        :alt: Cell with dash-dot colored border
        :figclass: align-center

        Cell with dash-dot colored border.


Apply Shadow to cell
""""""""""""""""""""

Using the :py:class:`~ooodev.format.calc.direct.cell.borders.Shadow` class shadows with a variety of options can be added to a cell.

In this example the default shadow is used.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        borders = Borders(border_side=Side(color=CommonColor.BLUE), shadow=Shadow())
        Styler.apply(cell, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210104021-d272159c-141a-4925-9232-e5b7a9594b8a:
    .. figure:: https://user-images.githubusercontent.com/4193389/210104021-d272159c-141a-4925-9232-e5b7a9594b8a.png
        :alt: Cell with blue colored border and default shadow
        :figclass: align-center

        Cell with blue colored border and default shadow.

Apply padding to a cell
"""""""""""""""""""""""

Using the :py:class:`~ooodev.format.calc.direct.cell.borders.Padding` class it is possible to add padding to a cell.
``Padding`` can take ``left``, ``right``, ``top``,  ``bottom`` arguments or ``all`` can be use to apply even padding to all sides at one.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        borders = Borders(border_side=Side(color=CommonColor.BLUE), padding=Padding(left=1.5))
        Styler.apply(cell, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210103438-0ddd7fa1-fd56-4caa-8d2b-209bf609adca:
    .. figure:: https://user-images.githubusercontent.com/4193389/210103438-0ddd7fa1-fd56-4caa-8d2b-209bf609adca.png
        :alt: Cell with blue colored border and left padding
        :figclass: align-center

        Cell with blue colored border and left padding.

.. cssclass:: screen_shot

    .. _230247760-76f6c21a-5dc8-476d-a4e7-9e6a8b6582ae:
    .. figure:: https://user-images.githubusercontent.com/4193389/230247760-76f6c21a-5dc8-476d-a4e7-9e6a8b6582ae.png
        :alt: Calc Format Cells Borders dialog
        :width: 450px
        :figclass: align-center

        Calc Format Cells Borders dialog

Cumulative borders
""""""""""""""""""

Applying more then one border style to a cell keeps previous formatting.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_obj = Calc.get_cell_obj("B2")
        Calc.set_val(value="Hello World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)

        border = Borders(diagonal_up=Side(color=CommonColor.RED))
        Styler.apply(cell, border)

        borders = Borders(diagonal_down=Side(color=CommonColor.BLUE))
        Styler.apply(cell, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210104021-9a796bf4-75c5-4867-a4ad-10331380905e:
    .. figure:: https://user-images.githubusercontent.com/4193389/210105163-9a796bf4-75c5-4867-a4ad-10331380905e.png
        :alt: Cell with cumulative borders
        :figclass: align-center

        Cell with cumulative borders.

Range of Cells
^^^^^^^^^^^^^^


Default Borders
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        Styler.apply(cr, Borders().default)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210106009-07a937e5-7d58-4329-85cf-e4e603f3e6f2:
    .. figure:: https://user-images.githubusercontent.com/4193389/210106009-07a937e5-7d58-4329-85cf-e4e603f3e6f2.png
        :alt: Range with default borders
        :figclass: align-center

        Range with default borders.


Removing Borders
""""""""""""""""

Appling ``Border().empty`` to a cell or a range will clear all formatting.

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        Styler.apply(cr, Borders().default)
        # ...
        Styler.apply(cr, Borders().empty)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Colored borders
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        borders = Borders(border_side=Side(color=CommonColor.RED))
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210106009-491db633-187c-42b7-a4ed-5ddd9e8a4a1e:
    .. figure:: https://user-images.githubusercontent.com/4193389/210112658-491db633-187c-42b7-a4ed-5ddd9e8a4a1e.png
        :alt: Range with colored borders
        :figclass: align-center

        Range with colored borders.

Applying border to a side
"""""""""""""""""""""""""

Apply green border to left side.

:py:class:`~ooodev.format.calc.direct.cell.borders.Borders` constructor can also take ``left``, ``right``, ``top``, ``bottom``, ``vertical``, ``horizontal``, ``diagonal_down`` and ``diagonal_up`` arguments as sides.
In this case just pass in the ``left`` side.

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        borders = Borders(left=Side(color=CommonColor.GREEN))
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210112804-00e54008-f2de-42d9-8a72-8ef7000c2b18:
    .. figure:: https://user-images.githubusercontent.com/4193389/210112804-00e54008-f2de-42d9-8a72-8ef7000c2b18.png
        :alt: Range with left colored border
        :figclass: align-center

        Range with left colored border.


Apply border with increased size
""""""""""""""""""""""""""""""""

Passing `width` argument to `Side()` controls border width.

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        side_left_right = Side(color=CommonColor.GREEN, width=1.8)
        borders = Borders(left=side_left_right, right=side_left_right)
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210112958-d25f44c0-75c5-49ef-bcaa-405f337e7878:
    .. figure:: https://user-images.githubusercontent.com/4193389/210112958-d25f44c0-5c9c-49ef-bcaa-405f337e7878.png
        :alt: Range with left and right colored border with increased width
        :figclass: align-center

        Range with left and right colored border with increased width.

Apply different top and bottom colors
"""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        side_top_bottom = Side(color=CommonColor.CHARTREUSE, width=1.3)
        side_left_right = Side(color=CommonColor.ROYAL_BLUE, width=1.3)
        borders = Borders(
            top=side_top_bottom,
            bottom=side_top_bottom,
            left=side_left_right,
            right=side_left_right,
        )
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113089-7c1e7a7e-2c1e-4645-a39f-5e2c80e4da0d:
    .. figure:: https://user-images.githubusercontent.com/4193389/210113089-cb1e7a7e-2c1e-4645-a39f-5e2c80e4da0d.png
        :alt: Range different top and bottom border colors
        :figclass: align-center

        Range different top and bottom border colors.


Apply Diagonal border
"""""""""""""""""""""

**UP**

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        borders = Borders(border_side=Side(), diagonal_up=Side(color=CommonColor.RED))
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113314-f656de46-4273-a786-5c445d00fe1b:
    .. figure:: https://user-images.githubusercontent.com/4193389/210113314-f656de46-8fc6-4273-a786-5c445d00fe1b.png
        :alt: Range with diagonal up border
        :figclass: align-center

        Range with diagonal up border.


**DOWN**

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        borders = Borders(border_side=Side(), diagonal_down=Side(color=CommonColor.RED))
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113401-1bca1147-76da-4df1-aabb-3f2cb856d66e:
    .. figure:: https://user-images.githubusercontent.com/4193389/210113401-1bca1147-76da-4df1-aabb-3f2cb856d66e.png
        :alt: Range with diagonal up border
        :figclass: align-center

        Range with diagonal up border.


Applying different style border
"""""""""""""""""""""""""""""""

Using py:class:`~ooodev.format.inner.direct.structs.side.BorderLineKind` enumeration it is possible to change the border style to many different configurations.

In this example the border style is set to Dash-dot.

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        borders = Borders(
            border_side=Side(line=BorderLineKind.DASH_DOT, color=CommonColor.DARK_GREEN)
        )
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113504-7ea66848-9e8e-4048-9d3c-c7a3ef20d7d6:
    .. figure:: https://user-images.githubusercontent.com/4193389/210113504-7ea66848-c9e8-4048-9d3c-c7a3ef20d7d6.png
        :alt: Range with dash-dot border
        :figclass: align-center

        Range with dash-dot border.

Apply Shadow to range
"""""""""""""""""""""

Using the :py:class:`~ooodev.format.calc.direct.cell.borders.Shadow` class shadows with a variety of options can be added to a range.

In this example the default shadow is used.

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        borders = Borders(border_side=Side(color=CommonColor.BLUE), shadow=Shadow())
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113632-e69f8bb2-484b-42e2-bfd6-508195f78cf0:
    .. figure:: https://user-images.githubusercontent.com/4193389/210113632-e69f8bb2-484b-42e2-bfd6-508195f78cf0.png
        :alt: Range with border and shadow
        :figclass: align-center

        Range with border and shadow.

Apply horizontal and vertical lines
"""""""""""""""""""""""""""""""""""

This example sets outer border to blue for all sides.
Horizontal lines are set to dash-dash-dot with color of green.
Vertical lines are set to double with a color of red.

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        borders = Borders(
            border_side=Side(color=CommonColor.BLUE),
            horizontal=Side(line=BorderLineKind.DASH_DOT_DOT, color=CommonColor.GREEN),
            vertical=Side(line=BorderLineKind.DOUBLE, color=CommonColor.RED),
        )
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113923-b916b3df-491c-4a9f-1-949e550fc138:
    .. figure:: https://user-images.githubusercontent.com/4193389/210113923-b916b3df-d443-491c-a9f1-949e550fc138.png
        :alt: Range various border styles
        :figclass: align-center

        Range various border styles.


Multiple Styles
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        rng_obj = Calc.get_range_obj("B2:F6")
        cr = Calc.get_cell_range(sheet, rng_obj)
        borders = Borders(
            border_side=Side(color=CommonColor.BLUE_VIOLET, width=1.3),
            diagonal_up=Side(color=CommonColor.RED),
            diagonal_down=Side(color=CommonColor.RED),
        )
        Styler.apply(cr, borders)

        rng_obj = Calc.get_range_obj("C3:E5")
        cr = Calc.get_cell_range(sheet, rng_obj)

        Styler.apply(cr, borders)

        borders = Borders(
            border_side=Side(color=CommonColor.BLUE),
            horizontal=Side(line=BorderLineKind.DASH_DOT_DOT, color=CommonColor.GREEN),
            vertical=Side(line=BorderLineKind.DOUBLE, color=CommonColor.RED),
        )
        Styler.apply(cr, borders)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210114562-c0d096c7-74c5-4905-a822-e2e123a7c1a0:
    .. figure:: https://user-images.githubusercontent.com/4193389/210114562-c0d096c6-f74c-4905-a822-e2e123a7c1a0.png
        :alt: Range multiple border styles
        :figclass: align-center

        Range multiple border styles.


.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_para_borders`