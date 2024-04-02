.. _help_chart2_format_direct_title_font_only:

Chart2 Direct Title/Subtitle Font Only
======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_font()`` method is used to set the font of the Title and Subtitle of a Chart,
similar options
as :numref:`0c6b3393-b925-42c1-9bc5-8b604459de9a` Font Dialog, but without the dialog.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 39,40,41,42,43

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

                title.style_font(
                    name="Lucida Calligraphy",
                    size=14,
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


Apply the Font
--------------

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_font(
            name="Lucida Calligraphy",
            size=14,
            font_style="italic",
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`0bd83e10-35ea-4ba3-bff9-04548d2ad0e0_1` and :numref:`0c6b3393-b925-42c1-9bc5-8b604459de9a_1`.

.. cssclass:: screen_shot

    .. _0bd83e10-35ea-4ba3-bff9-04548d2ad0e0_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/0bd83e10-35ea-4ba3-bff9-04548d2ad0e0
        :alt: Chart with Title Font set
        :figclass: align-center
        :width: 450px

        Chart with Title Font set


.. cssclass:: screen_shot

    .. _0c6b3393-b925-42c1-9bc5-8b604459de9a_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/0c6b3393-b925-42c1-9bc5-8b604459de9a
        :alt: Chart Data Labels Dialog Font
        :figclass: align-center
        :width: 450px

        Chart Data Labels Dialog Font

Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_font(
            name="Lucida Calligraphy",
            size=14,
            font_style="italic",
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`6427af0a-2fad-4f6a-b390-813c9503eced_1`.

.. cssclass:: screen_shot

    .. _6427af0a-2fad-4f6a-b390-813c9503eced_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6427af0a-2fad-4f6a-b390-813c9503eced
        :alt: Chart with Subtitle Font set
        :figclass: align-center
        :width: 450px

        Chart with Subtitle Font set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_title_font_effects`
        - :ref:`help_chart2_format_direct_title_font`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`