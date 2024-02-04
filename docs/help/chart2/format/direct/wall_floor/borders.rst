.. _help_chart2_format_direct_wall_floor_borders:

Chart2 Direct Wall/Floor Borders
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.wall.borders.LineProperties` class gives the same options as the Chart Wall Borders dialog
as seen in :numref:`b608a83b-aa0a-4389-aee0-6d65a2b5536e`.


.. cssclass:: screen_shot

    .. _b608a83b-aa0a-4389-aee0-6d65a2b5536e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b608a83b-aa0a-4389-aee0-6d65a2b5536e
        :alt: Chart Area Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog

Setup
-----

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.chart2.direct.wall.borders import LineProperties as WallLineProperties
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("col_chart3d.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.BLUE_DARK1, width=1.0)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_bdr_line])

                wall_bdr_line = WallLineProperties(
                    color=StandardColor.PURPLE_DARK2, width=1.2, transparency=30
                )
                Chart2.style_wall(chart_doc=chart_doc, styles=[wall_bdr_line])

                floor_bdr_line = WallLineProperties(
                    color=StandardColor.PURPLE_DARK1, width=0.8, transparency=20
                )
                Chart2.style_floor(chart_doc=chart_doc, styles=[floor_bdr_line])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting Line Properties
-----------------------

The :py:class:`~ooodev.format.chart2.direct.wall.borders.LineProperties` class is used to set the border line properties.

Before applying formatting is seen in :numref:`fceab75a-31d7-4742-a331-83a79232b783`.

Apply to wall.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.wall.borders import LineProperties as WallLineProperties
        # ... other code

        wall_bdr_line = WallLineProperties(color=StandardColor.PURPLE_DARK2, width=1.2, transparency=30)
        Chart2.style_wall(chart_doc=chart_doc, styles=[wall_bdr_line])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to floor.

.. tabs::

    .. code-tab:: python

        floor_bdr_line = WallLineProperties(color=StandardColor.PURPLE_DARK1, width=0.8, transparency=20)
        Chart2.style_floor(chart_doc=chart_doc, styles=[floor_bdr_line])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`36619546-8646-4d90-90eb-ed1cac18f986` and :numref:`2568b994-2e62-4401-bb8d-29c3b07a005e`


.. cssclass:: screen_shot

    .. _36619546-8646-4d90-90eb-ed1cac18f986:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/36619546-8646-4d90-90eb-ed1cac18f986
        :alt: Chart with wall and floor border set
        :figclass: align-center
        :width: 450px

        Chart with wall and floor border set

.. cssclass:: screen_shot

    .. _2568b994-2e62-4401-bb8d-29c3b07a005e:

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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_wall() <ooodev.office.chart2.Chart2.style_wall>`
        - :py:meth:`Chart2.style_floor() <ooodev.office.chart2.Chart2.style_floor>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.wall.borders.LineProperties`