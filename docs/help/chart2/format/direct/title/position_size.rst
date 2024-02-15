.. _help_chart2_format_direct_title_position_size:

Chart2 Direct Title/Subtitle Position
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_position()`` method is called to set the position of the chart title.

Setup
-----

General setup used to run the examples in this page.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 39

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.utils.data_type.color_range import ColorRange

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "piechart.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.PURPLE_DARK1,
                    width=0.7,
                )
                _ = chart_doc.style_area_gradient(
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(
                        StandardColor.BLUE_DARK1,
                        StandardColor.PURPLE_LIGHT2,
                    ),
                )

                title = chart_doc.get_title()
                if title is None:
                    raise ValueError("Title not found")

                title.style_position(7.1, 66.3)

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Applying Position
-----------------

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

By default the :py:class:`~ooodev.format.chart2.direct.title.position_size.Position` class uses millimeters as the unit of measure.

Apply to Title
^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_position(7.1, 66.3)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`3c13137c-0b86-47b5-9b34-ee52902aff0f_1` and :numref:`e92ab05a-6093-43ce-a83b-14862827ec35_1`.

.. cssclass:: screen_shot

    .. _3c13137c-0b86-47b5-9b34-ee52902aff0f_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3c13137c-0b86-47b5-9b34-ee52902aff0f
        :alt: Chart with title position set
        :figclass: align-center
        :width: 450px

        Chart with title position set

.. cssclass:: screen_shot

    .. _bfd22d03-f4d8-4d1e-9759-b773051c79df_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bfd22d03-f4d8-4d1e-9759-b773051c79df
        :alt: Chart Title Position and Size Dialog
        :figclass: align-center

        Chart Title Position and Size Dialog

Apply to Subtitle
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_position(7.1, 66.3)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`3ee5c63f-f82e-4958-9d6c-cde4eaaf3f4f_1`.

.. cssclass:: screen_shot

    .. _3ee5c63f-f82e-4958-9d6c-cde4eaaf3f4f_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3ee5c63f-f82e-4958-9d6c-cde4eaaf3f4f
        :alt: Chart with subtitle position set
        :figclass: align-center
        :width: 450px

        Chart with subtitle position set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_title`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
