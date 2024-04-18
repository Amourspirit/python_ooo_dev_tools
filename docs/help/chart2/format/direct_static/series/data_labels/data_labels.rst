.. _help_chart2_format_direct_static_series_labels_data_labels:

Chart2 Direct Series Data Labels (Static)
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

The Data Labels tab of the Chart Data Labels dialog has many options as see in :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.
The :py:mod:`ooodev.format.chart2.direct.series.data_labels.data_labels` module has various classes to set the same options.

Calls to the :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
and :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>` methods are used to set the data labels options of a Chart.

.. cssclass:: screen_shot

    .. _0c7d3398-34f5-4da3-81f2-79e134fab44d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/0c7d3398-34f5-4da3-81f2-79e134fab44d
        :alt: Chart Data Labels Dialog Data Labels Tab
        :figclass: align-center
        :width: 450px

        Chart Data Labels Dialog Data Labels Tab

.. seealso::

    - :ref:`help_chart2_format_direct_series_labels_data_labels`

Setup
-----

General setup for these examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 33,33,34,35,36,37,38,39

        import uno
        from ooodev.format.chart2.direct.series.data_labels.data_labels import PercentFormat
        from ooodev.format.chart2.direct.series.data_labels.data_labels import TextAttribs
        from ooodev.format.chart2.direct.series.data_labels.data_labels import NumberFormatIndexEnum
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient, PresetGradientKind
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.wall.transparency import Transparency as WallTransparency
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("col_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_LIGHT3, width=0.7)
                chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.NEON_LIGHT)
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                wall_transparency = WallTransparency(value=60)
                Chart2.style_wall(chart_doc=chart_doc, styles=[wall_transparency])

                text_attribs = TextAttribs(show_number=True)
                format_number = NumberFormat(
                    chart_doc=chart_doc,
                    source_format=False,
                    num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2,
                )
                Chart2.style_data_series(chart_doc=chart_doc, styles=[text_attribs, format_number])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Text Attributes
---------------

The text attributes are set using three classes that are covered in this section.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Text Attribs
""""""""""""

The :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.TextAttribs` class is used to set the various boolean options in the ``Text Attributes`` section of the Chart Data Labels dialog as seen in :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_labels.data_labels import TextAttribs
        # ... other code

        text_attribs = TextAttribs(
            show_category_name=True,
            show_legend_symbol=True,
            show_series_name=True,
            auto_text_wrap=True,
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[text_attribs])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`ffd2621d-fb71-4a00-ad8a-5d0760ed11bf` and :numref:`7852b8b7-054a-484c-823c-3512b700132b`.

.. cssclass:: screen_shot

    .. _ffd2621d-fb71-4a00-ad8a-5d0760ed11bf:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ffd2621d-fb71-4a00-ad8a-5d0760ed11bf
        :alt: Chart with formatting applied to data series
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data series

.. cssclass:: screen_shot

    .. _7852b8b7-054a-484c-823c-3512b700132b:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7852b8b7-054a-484c-823c-3512b700132b
        :alt: Chart Format Number Dialog
        :figclass: align-center
        :width: 450px

        Chart Format Number Dialog

Style Data Point
~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=2, styles=[text_attribs])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`959761b7-4336-4712-8e86-a48897606925`.

.. cssclass:: screen_shot

    .. _959761b7-4336-4712-8e86-a48897606925:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/959761b7-4336-4712-8e86-a48897606925
        :alt: Chart with Text Attributes applied to data point
        :figclass: align-center
        :width: 450px

        Chart with Text Attributes applied to data point

Number Format
"""""""""""""

The :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.NumberFormat` class is used to set the number format of the data labels.
This class is used to set the values seen in :numref:`ca21f3f1-e1b1-4bab-bb36-f52c966e00af`.

The ``NumberFormatIndexEnum`` enum contains the values in |num_fmt_index|_ for easy lookup.

To ensure that the effects of :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.NumberFormat` are
visible the :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.TextAttribs` class is used to
turn on ``Value as Number`` of the dialog seen in :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_labels.data_labels import TextAttribs
        from ooodev.format.chart2.direct.series.data_labels.data_labels import NumberFormat
        from ooodev.format.chart2.direct.series.data_labels.data_labels import NumberFormatIndexEnum
        # ... other code

        text_attribs = TextAttribs(show_number=True)
        format_number = NumberFormat(
            chart_doc=chart_doc,
            source_format=False,
            num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2,
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[text_attribs, format_number])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`3d1f582b-558d-4da5-8996-bebb6b6781d0` and :numref:`ca21f3f1-e1b1-4bab-bb36-f52c966e00af`.

.. cssclass:: screen_shot

    .. _3d1f582b-558d-4da5-8996-bebb6b6781d0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3d1f582b-558d-4da5-8996-bebb6b6781d0
        :alt: Chart with Text Attributes applied to data series
        :figclass: align-center
        :width: 450px

        Chart with Text Attributes applied to data series

.. cssclass:: screen_shot

    .. _ca21f3f1-e1b1-4bab-bb36-f52c966e00af:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ca21f3f1-e1b1-4bab-bb36-f52c966e00af
        :alt: Chart Format Number Dialog
        :figclass: align-center
        :width: 450px

        Chart Format Number Dialog

Style Data Point
~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(
            chart_doc=chart_doc, series_idx=0, idx=1, styles=[text_attribs, format_number]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`157ea466-4511-4f84-90e9-52b76390c1fb`.

.. cssclass:: screen_shot

    .. _157ea466-4511-4f84-90e9-52b76390c1fb:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/157ea466-4511-4f84-90e9-52b76390c1fb
        :alt: Chart with Text Attributes applied to data point
        :figclass: align-center
        :width: 450px

        Chart with Text Attributes applied to data point

Percentage Format
"""""""""""""""""

The :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.PercentFormat` class is used to set the number format of the data labels.
This class is used to set the values seen in :numref:`ca21f3f1-e1b1-4bab-bb36-f52c966e00af`.

The ``NumberFormatIndexEnum`` enum contains the values in |num_fmt_index|_ for easy lookup.

To ensure that the effects of :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.PercentFormat` are
visible the :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.TextAttribs` class is used to
turn on ``Value as Percentage`` of the dialog seen in :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_labels.data_labels import PercentFormat
        from ooodev.format.chart2.direct.series.data_labels.data_labels import NumberFormatIndexEnum
        # ... other code

        text_attribs = TextAttribs(show_number_in_percent=True)
        format_percent = PercentFormat(
            chart_doc=chart_doc,
            source_format=False,
            num_format_index=NumberFormatIndexEnum.PERCENT_DEC2,
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[text_attribs, format_percent])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`d8b1329b-d94e-457d-91d1-87d5f14aefa2` and :numref:`45c0d0a1-4c9e-4b84-ad9b-c92bb4a2658e`.

.. cssclass:: screen_shot

    .. _d8b1329b-d94e-457d-91d1-87d5f14aefa2:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d8b1329b-d94e-457d-91d1-87d5f14aefa2
        :alt: Chart with formatting applied to data series
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data series

.. cssclass:: screen_shot

    .. _45c0d0a1-4c9e-4b84-ad9b-c92bb4a2658e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/45c0d0a1-4c9e-4b84-ad9b-c92bb4a2658e
        :alt: Chart Format Number Dialog
        :figclass: align-center
        :width: 450px

        Chart Format Number Dialog

Style Data Point
~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_data_point(
            chart_doc=chart_doc, series_idx=0, idx=3, styles=[text_attribs, format_percent]
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`cc247b93-54e7-4f51-a5c7-c80c759eaad8`.

.. cssclass:: screen_shot

    .. _cc247b93-54e7-4f51-a5c7-c80c759eaad8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/cc247b93-54e7-4f51-a5c7-c80c759eaad8
        :alt: Chart with formatting applied to data point
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data point

Attribute Options
-----------------

The :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.AttribOptions` class is used to set the Options data labels.
This class is used to set the values seen in the ``Attribute Options`` section of :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

The :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.PlacementKind` enum is used to look up the placement.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_labels.data_labels import AttribOptions
        from ooodev.format.chart2.direct.series.data_labels.data_labels import PlacementKind
        # ... other code

        attrib_opt = AttribOptions(placement=PlacementKind.INSIDE)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[attrib_opt])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`115e2eaa-876c-4048-b30a-06e5be91b240` and :numref:`6b9458d9-b457-4de2-aa54-7c44a711e2a2`.

.. cssclass:: screen_shot

    .. _115e2eaa-876c-4048-b30a-06e5be91b240:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/115e2eaa-876c-4048-b30a-06e5be91b240
        :alt: Chart with formatting applied to data series
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data series

.. cssclass:: screen_shot

    .. _6b9458d9-b457-4de2-aa54-7c44a711e2a2:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6b9458d9-b457-4de2-aa54-7c44a711e2a2
        :alt: Chart Format Number Dialog
        :figclass: align-center
        :width: 450px

        Chart Format Number Dialog

Style Data Point
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[attrib_opt])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`4968c491-5e45-449e-800f-01549bc009bd`.

.. cssclass:: screen_shot

    .. _4968c491-5e45-449e-800f-01549bc009bd:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4968c491-5e45-449e-800f-01549bc009bd
        :alt: Chart with formatting applied to data point
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data point

Rotate Text
-----------

The :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.Orientation` class is used to set the rotation of data labels.
This class is used to set the values seen in the ``Rotate Text`` section of :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

The :py:class:`~ooodev.format.inner.direct.chart2.title.alignment.direction.DirectionModeKind` enum is used to look up the text direction.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.series.data_labels.data_labels import Orientation
        from ooodev.format.chart2.direct.series.data_labels.data_labels import DirectionModeKind
        # ... other code

        rotation = Orientation(angle=60, mode=DirectionModeKind.LR_TB, leaders=True)
        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[rotation])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`d57bc634-0f1e-4acc-9d02-848809635021` and :numref:`91cac9f6-9dbb-4017-a682-cd7a977c208e`.

.. cssclass:: screen_shot

    .. _d57bc634-0f1e-4acc-9d02-848809635021:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d57bc634-0f1e-4acc-9d02-848809635021
        :alt: Chart with formatting applied to data series
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data series

.. cssclass:: screen_shot

    .. _91cac9f6-9dbb-4017-a682-cd7a977c208e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/91cac9f6-9dbb-4017-a682-cd7a977c208e
        :alt: Chart Format Number Dialog
        :figclass: align-center
        :width: 450px

        Chart Format Number Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=2, styles=[rotation])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`35ff95c1-f3b3-48d6-848f-8c2935faa9b3`

.. cssclass:: screen_shot

    .. _35ff95c1-f3b3-48d6-848f-8c2935faa9b3:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/35ff95c1-f3b3-48d6-848f-8c2935faa9b3
        :alt: Chart with formatting applied to data point
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data point


Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_chart2_format_direct_series_labels_data_labels`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - |num_fmt|_
        - |num_fmt_index|_
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_data_series() <ooodev.office.chart2.Chart2.style_data_series>`
        - :py:meth:`Chart2.style_data_point() <ooodev.office.chart2.Chart2.style_data_point>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.series.data_labels.data_labels.TextAttribs`
        - :py:class:`ooodev.format.chart2.direct.series.data_labels.data_labels.NumberFormat`
        - :py:class:`ooodev.format.chart2.direct.series.data_labels.data_labels.PercentFormat`
        - :py:class:`ooodev.format.chart2.direct.series.data_labels.data_labels.AttribOptions`
        - :py:class:`ooodev.format.chart2.direct.series.data_labels.data_labels.Orientation`

.. |num_fmt| replace:: API NumberFormat
.. _num_fmt: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html

.. |num_fmt_index| replace:: API NumberFormatIndex
.. _num_fmt_index: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html
