.. _help_chart2_format_direct_general_transparency:

Chart2 Direct General Transparency
==================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_area_transparency_*()`` method can be used to set chart transparency.

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from pathlib import Path
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.utils.color import StandardColor
        from ooodev.loader.lo import Lo


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "col_chart.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_area_color(color=StandardColor.GREEN_LIGHT2)
                _ = chart_doc.style_area_transparency_transparency(50)

                Lo.delay(1_000)
                doc.close()
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Transparency
------------

Before setting the background transparency of the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Setting Transparency
^^^^^^^^^^^^^^^^^^^^

The ``style_area_transparency_transparency()`` method can be used to set the transparency of a chart.

.. tabs::

    .. code-tab:: python

        _ = chart_doc.style_area_color(color=StandardColor.GREEN_LIGHT2)
        _ = chart_doc.style_area_transparency_transparency(50)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`236953627-38d0f2c5-e19a-402e-942c-d7f1c1a27c11_1` and :numref:`236953723-96edce28-2476-4abb-af3d-223723c4fd1a_1`.

.. cssclass:: screen_shot

    .. _236953627-38d0f2c5-e19a-402e-942c-d7f1c1a27c11_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236953627-38d0f2c5-e19a-402e-942c-d7f1c1a27c11.png
        :alt: Chart with border, color and  transparency
        :figclass: align-center
        :width: 450px

        Chart with border, color and  transparency

.. cssclass:: screen_shot

    .. _236953723-96edce28-2476-4abb-af3d-223723c4fd1a_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236953723-96edce28-2476-4abb-af3d-223723c4fd1a.png
        :alt: Chart Area Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Transparency Dialog

Getting the transparency from a Chart
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = chart_doc.style_area_transparency_transparency_get()
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Gradient Transparency
---------------------

Before setting the background gradient transparency of the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Setting Gradient
^^^^^^^^^^^^^^^^

The ``style_area_transparency_gradient()`` method can be used to set the gradient transparency of a chart.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.data_type.intensity_range import IntensityRange
        # ... other code

        _ = chart_doc.style_area_transparency_gradient(angle=30, grad_intensity=IntensityRange(0, 100))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results can bee seen in :numref:`236955053-0dba0e1e-6bbf-4b22-921b-5e19e2131baf_1` and :numref:`236955121-cad9d1e7-c86d-435f-920c-02e0bb451c84_1`.

.. cssclass:: screen_shot

    .. _236955053-0dba0e1e-6bbf-4b22-921b-5e19e2131baf_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236955053-0dba0e1e-6bbf-4b22-921b-5e19e2131baf.png
        :alt: Chart with border, color and  transparency
        :figclass: align-center
        :width: 450px

        Chart with border, color and  transparency

.. cssclass:: screen_shot

    .. _236955121-cad9d1e7-c86d-435f-920c-02e0bb451c84_1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236955121-cad9d1e7-c86d-435f-920c-02e0bb451c84.png
        :alt: Chart Area Transparency Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Transparency Dialog

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - :ref:`help_chart2_format_direct_wall_floor_transparency`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.general.transparency.Transparency`
        - :py:class:`ooodev.format.chart2.direct.general.transparency.Gradient`
