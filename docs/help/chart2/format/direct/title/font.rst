.. _help_chart2_format_direct_title_font:

Chart2 Direct Title/Subtitle Font
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_font_general()`` method is used to set the font properties of a Chart Title or Subtitle.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 39,40,41,42,43,44,45

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

                title.style_font_general(
                    b=True,
                    i=True,
                    u=True,
                    color=StandardColor.PURPLE_DARK2,
                    shadowed=True,
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

        from ooodev.utils.color import StandardColor

        # ... other code
        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_font_general(
            b=True,
            i=True,
            u=True,
            color=StandardColor.PURPLE_DARK2,
            shadowed=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`eaa1eab4-687c-466a-a7fd-2c126f7b1e2f_1`.

.. cssclass:: screen_shot

    .. _eaa1eab4-687c-466a-a7fd-2c126f7b1e2f_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/eaa1eab4-687c-466a-a7fd-2c126f7b1e2f
        :alt: Chart with Title Font set
        :figclass: align-center
        :width: 450px

        Chart with Title Font set


Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.utils.color import StandardColor

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_font_general(
            b=True,
            i=True,
            u=True,
            color=StandardColor.PURPLE_DARK2,
            shadowed=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None



Running the above code will produce the following output shown in :numref:`bb19afad-c492-4f6f-a7bf-89d2323b1c77_1`.

.. cssclass:: screen_shot

    .. _bb19afad-c492-4f6f-a7bf-89d2323b1c77_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bb19afad-c492-4f6f-a7bf-89d2323b1c77
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
        - :ref:`help_chart2_format_direct_title_font_only`
        - :ref:`help_chart2_format_direct_title_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`