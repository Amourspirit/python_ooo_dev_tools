.. _help_calc_format_direct_cell_borders:

Calc Direct Cell Borders
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 3

Overview
--------

Calc has a dialog, as seen in :numref:`ss_calc_format_cell_borders_dialog_1`, that sets cell borders. In this section we will look the various classes that set the same options.

The :py:class:`ooodev.format.calc.direct.cell.borders.Borders` class is used to set the border values.

.. cssclass:: screen_shot

    .. _ss_calc_format_cell_borders_dialog_1:

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

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.format.calc.direct.cell.borders import Side
        from ooodev.utils.color import CommonColor


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(130)

                rng = sheet.rng("B2:F6")
                cell_rng = sheet.get_range(range_obj=rng)
                cell_rng.style_borders(border_side=Side(color=CommonColor.BLUE))

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

Single Cell
^^^^^^^^^^^

Default Border
""""""""""""""

Calling ``style_borders_default()`` will create a default border for a cell or a range.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        cell.style_borders_default()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210101040-aa66cae1-323b-4fb0-b9c2-ac3a82a62403_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210101040-aa66cae1-323b-4fb0-b9c2-ac3a82a62403.png
        :alt: Cell with default border
        :figclass: align-center

        Cell with default border.

Removing Border
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        cell.style_borders_default()
        # ...
        # remove border
        cell.style_borders_clear()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Colored border
""""""""""""""

The ``style_borders_sides()`` method can be used when only all four sides are to be styled at once.
Here The ``style_borders_sides()`` is used to set the border color to red.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        cell.style_borders_sides(color=CommonColor.RED)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210101175-74a38aa2-c77a-4f6c-ad76-3b3f2322c6d9_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210101175-74a38aa2-c77a-4f6c-ad76-3b3f2322c6d9.png
        :alt: Cell with colored border
        :figclass: align-center

        Cell with colored border.

Applying border to a side
"""""""""""""""""""""""""

Apply green border to left side.

The ``style_borders()`` method can also take ``left``, ``right``, ``top``, ``bottom``, ``vertical``, ``horizontal``, ``diagonal_down`` and ``diagonal_up`` arguments as sides.
In this case just pass in the ``left`` side.

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell.style_borders(
            left=Side(color=CommonColor.GREEN),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210101363-4288e162-2117-4b95-bed0-578a179b31f1_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210101363-4288e162-2117-4b95-bed0-578a179b31f1.png
        :alt: Cell with left colored border
        :figclass: align-center

        Cell with left colored border.

Apply border with increased size
""""""""""""""""""""""""""""""""

Passing ``width`` argument to ``Side()`` controls border width.

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        side_left_right = Side(color=CommonColor.GREEN, width=1.8)
        cell.style_borders(
            left=side_left_right, right=side_left_right
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210101564-b76cd842-ed82-4fd9-85b6-16890cb80364_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210100564-b76cd842-ed82-4fd9-85b6-16890cb80364.png
        :alt: Cell with left and right colored border
        :figclass: align-center

        Cell with left and right colored border.

Apply different top and side colors
"""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        side_top_bottom = Side(color=CommonColor.CHARTREUSE, width=1.3)
        side_left_right = Side(color=CommonColor.ROYAL_BLUE, width=1.3)
        cell.style_borders(
            top=side_top_bottom,
            bottom=side_top_bottom,
            left=side_left_right,
            right=side_left_right,
        )

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

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        cell.style_borders(
            diagonal_up=Side(color=CommonColor.RED)
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210102706-ebe79c91-5e0a-4482-a58d-a797efa7ded9_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210102706-ebe79c91-5e0a-4482-a58d-a797efa7ded9.png
        :alt: Cell with diagonal up colored border
        :figclass: align-center

        Cell with diagonal up colored border.


**DOWN**

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        cell.style_borders(
            diagonal_down=Side(color=CommonColor.RED)
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210102945-73b453d6-33f2-4582-a276-61fda1e5edbe_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210102945-73b453d6-33f2-4582-a276-61fda1e5edbe.png
        :alt: Cell with diagonal down colored border
        :figclass: align-center

        Cell with diagonal down colored border.

Applying different style border
"""""""""""""""""""""""""""""""

Using :py:class:`~ooodev.format.inner.direct.structs.side.BorderLineKind` enumeration it is possible to change the border style to many different configurations.

In this example the border style is set to Dash-dot.

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import BorderLineKind

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        cell.style_borders_sides(
            line=BorderLineKind.DASH_DOT,
            color=CommonColor.DARK_GREEN
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210103415-147a46c0-7d99-4cd4-b861-d46228a89c25_1:

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

        from ooodev.format.calc.direct.cell.borders import Shadow

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        cell.style_borders_sides(
            color=CommonColor.BLUE,
            shadow=Shadow(),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210104021-d272159c-141a-4925-9232-e5b7a9594b8a_1:

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

        from ooodev.format.calc.direct.cell.borders import Padding

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        cell.style_borders_sides(
            color=CommonColor.BLUE,
            padding=Padding(left=1.5),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210103438-0ddd7fa1-fd56-4caa-8d2b-209bf609adca_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210103438-0ddd7fa1-fd56-4caa-8d2b-209bf609adca.png
        :alt: Cell with blue colored border and left padding
        :figclass: align-center

        Cell with blue colored border and left padding.

.. cssclass:: screen_shot

    .. _230247760-76f6c21a-5dc8-476d-a4e7-9e6a8b6582ae_1:

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

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell = sheet["B2"]
        cell.value = "Hello World"
        cell.style_borders(diagonal_up=Side(color=CommonColor.RED))
        cell.style_borders(diagonal_down=Side(color=CommonColor.BLUE))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210104021-9a796bf4-75c5-4867-a4ad-10331380905e_1:

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
        cell_rng = sheet.get_range(range_name="B2:F6")
        cell_rng.style_borders_default()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210106009-07a937e5-7d58-4329-85cf-e4e603f3e6f2_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210106009-07a937e5-7d58-4329-85cf-e4e603f3e6f2.png
        :alt: Range with default borders
        :figclass: align-center

        Range with default borders.


Removing Borders
""""""""""""""""

Applying ``Border().empty`` to a cell or a range will clear all formatting.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_rng = sheet.get_range(range_name="B2:F6")
        cell_rng.style_borders_default()
        # ...
        cell_rng.style_borders_clear()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Colored borders
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_rng = sheet.get_range(range_name="B2:F6")
        cell_rng.style_borders_sides(color=CommonColor.RED)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210106009-491db633-187c-42b7-a4ed-5ddd9e8a4a1e_1:

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

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell_rng = sheet.get_range(range_name="B2:F6")
        cell_rng.style_borders(left=Side(color=CommonColor.GREEN))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210112804-00e54008-f2de-42d9-8a72-8ef7000c2b18_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210112804-00e54008-f2de-42d9-8a72-8ef7000c2b18.png
        :alt: Range with left colored border
        :figclass: align-center

        Range with left colored border.


Apply border with increased size
""""""""""""""""""""""""""""""""

Passing `width` argument to `Side()` controls border width.

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell_rng = sheet.get_range(range_name="B2:F6")
        side_left_right = Side(color=CommonColor.GREEN, width=1.8)
        cell_rng.style_borders(
            left=side_left_right, right=side_left_right
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210112958-d25f44c0-75c5-49ef-bcaa-405f337e7878_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210112958-d25f44c0-5c9c-49ef-bcaa-405f337e7878.png
        :alt: Range with left and right colored border with increased width
        :figclass: align-center

        Range with left and right colored border with increased width.

Apply different top and side colors
"""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell_rng = sheet.get_range(range_name="B2:F6")
        side_top_bottom = Side(color=CommonColor.CHARTREUSE, width=1.3)
        side_left_right = Side(color=CommonColor.ROYAL_BLUE, width=1.3)
        cell_rng.style_borders(
            top=side_top_bottom,
            bottom=side_top_bottom,
            left=side_left_right,
            right=side_left_right,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113089-7c1e7a7e-2c1e-4645-a39f-5e2c80e4da0d_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210113089-cb1e7a7e-2c1e-4645-a39f-5e2c80e4da0d.png
        :alt: Range different top and bottom border colors
        :figclass: align-center

        Range different top and bottom border colors.


Apply Diagonal border
"""""""""""""""""""""

**UP**

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell_rng = sheet.get_range(range_name="B2:F6")
        cell_rng.style_borders(
            border_side=Side(),
            diagonal_up=Side(color=CommonColor.RED),
        )


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113314-f656de46-4273-a786-5c445d00fe1b_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210113314-f656de46-8fc6-4273-a786-5c445d00fe1b.png
        :alt: Range with diagonal up border
        :figclass: align-center

        Range with diagonal up border.


**DOWN**

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import Side

        # ... other code
        cell_rng = sheet.get_range(range_name="B2:F6")
        cell_rng.style_borders(
            border_side=Side(),
            diagonal_down=Side(color=CommonColor.RED),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113401-1bca1147-76da-4df1-aabb-3f2cb856d66e_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210113401-1bca1147-76da-4df1-aabb-3f2cb856d66e.png
        :alt: Range with diagonal up border
        :figclass: align-center

        Range with diagonal up border.


Applying different style border
"""""""""""""""""""""""""""""""

Using :py:class:`~ooodev.format.inner.direct.structs.side.BorderLineKind` enumeration it is possible to change the border style to many different configurations.

In this example the border style is set to Dash-dot.

.. tabs::

    .. code-tab:: python

        # ... other code
        cell_rng = sheet.get_range(range_name="B2:F6")
        cell_rng.style_borders_sides(
            line=BorderLineKind.DASH_DOT,
            color=CommonColor.DARK_GREEN,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113504-7ea66848-9e8e-4048-9d3c-c7a3ef20d7d6_1:

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

        from ooodev.format.calc.direct.cell.borders import Shadow

        # ... other code
        cell_rng = sheet.get_range(range_name="B2:F6")
        cell_rng.style_borders_sides(
            color=CommonColor.BLUE,
            shadow=Shadow(),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113632-e69f8bb2-484b-42e2-bfd6-508195f78cf0_1:

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

        from ooodev.format.calc.direct.cell.borders import Side
        from ooodev.format.calc.direct.cell.borders import BorderLineKind

        # ... other code
        ell_rng = sheet.get_range(range_name="B2:F6")
        cell_rng.style_borders(
            border_side=Side(color=CommonColor.BLUE),
            horizontal=Side(line=BorderLineKind.DASH_DOT_DOT, color=CommonColor.GREEN),
            vertical=Side(line=BorderLineKind.DOUBLE, color=CommonColor.RED),
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210113923-b916b3df-491c-4a9f-1-949e550fc138_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210113923-b916b3df-d443-491c-a9f1-949e550fc138.png
        :alt: Range various border styles
        :figclass: align-center

        Range various border styles.


Multiple Styles
"""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.direct.cell.borders import Side
        from ooodev.format.calc.direct.cell.borders import BorderLineKind

        # ... other code
        with doc:
            # lock document controllers for fast processing and avoid flickering.
            cell_rng = sheet.get_range(range_name="B2:F6")
            cell_rng.style_borders(
                border_side=Side(color=CommonColor.BLUE_VIOLET, width=1.3),
                diagonal_up=Side(color=CommonColor.RED),
                diagonal_down=Side(color=CommonColor.RED),
            )

            cell_rng = sheet.get_range(range_name="C3:E5")
            cell_rng.style_borders_clear()
            cell_rng.style_borders(
                border_side=Side(color=CommonColor.BLUE),
                horizontal=Side(line=BorderLineKind.DASH_DOT_DOT, color=CommonColor.GREEN),
                vertical=Side(line=BorderLineKind.DOUBLE, color=CommonColor.RED),
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210114562-c0d096c7-74c5-4905-a822-e2e123a7c1a0_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/210114562-c0d096c6-f74c-4905-a822-e2e123a7c1a0.png
        :alt: Range multiple border styles
        :figclass: align-center

        Range multiple border styles.

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_para_borders`
        - :ref:`help_writer_format_direct_table`
        - :ref:`help_calc_format_modify_cell_borders`
        - :py:class:`ooodev.format.calc.direct.cell.borders.Borders`
        - :py:class:`ooodev.format.calc.direct.cell.borders.Padding`
        - :py:class:`ooodev.format.calc.direct.cell.borders.Shadow`