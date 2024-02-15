.. _help_chart2_format_direct_wall_floor_borders:

Chart2 Direct Wall/Floor Borders
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_border_line()`` method gives the same options as the Chart Wall Borders dialog
as seen in :numref:`b608a83b-aa0a-4389-aee0-6d65a2b5536e_1`.


.. cssclass:: screen_shot

    .. _b608a83b-aa0a-4389-aee0-6d65a2b5536e_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b608a83b-aa0a-4389-aee0-6d65a2b5536e
        :alt: Chart Area Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 25,26,27

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "col_chart3d.ods"
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

                wall = chart_doc.first_diagram.wall
                wall.style_border_line(
                    StandardColor.PURPLE_DARK1, width=0.8, transparency=20
                )

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting Line Properties
-----------------------

The ``style_border_line()`` method is called to set the border line properties.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.color import StandardColor

        # ... other code
        wall = chart_doc.first_diagram.wall
        wall.style_border_line(
            StandardColor.PURPLE_DARK1, width=0.8, transparency=20
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.color import StandardColor

        # ... other code
        floor = chart_doc.first_diagram.floor
        floor.style_border_line(
            StandardColor.PURPLE_DARK1, width=0.8, transparency=20
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`36619546-8646-4d90-90eb-ed1cac18f986_1` and :numref:`2568b994-2e62-4401-bb8d-29c3b07a005e_1`


.. cssclass:: screen_shot

    .. _36619546-8646-4d90-90eb-ed1cac18f986_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/36619546-8646-4d90-90eb-ed1cac18f986
        :alt: Chart with wall and floor border set
        :figclass: align-center
        :width: 450px

        Chart with wall and floor border set

.. cssclass:: screen_shot

    .. _2568b994-2e62-4401-bb8d-29c3b07a005e_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2568b994-2e62-4401-bb8d-29c3b07a005e
        :alt: Chart Area Borders Default Dialog Modified
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog Modified

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_general_borders`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`