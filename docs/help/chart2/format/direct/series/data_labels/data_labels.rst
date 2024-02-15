.. _help_chart2_format_direct_series_labels_data_labels:

Chart2 Direct Series Data Labels
================================

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

Setup
-----

General setup for these examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 32,33,34,35,36

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
        from ooodev.utils.color import StandardColor

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
                _ = chart_doc.style_border_line(
                    color=StandardColor.GREEN_LIGHT3,
                    width=0.7,
                )
                _ = chart_doc.style_area_gradient_from_preset(
                    preset=PresetGradientKind.NEON_LIGHT,
                )

                chart_doc.first_diagram.wall.style_area_transparency_transparency(60)
                ds = chart_doc.get_data_series()[0]

                ds.style_text_attributes(show_number=True)
                ds.style_numbers_numbers(
                    source_format=False,
                    num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2,
                )

                # f_style = ds.style_label_border_line_get()
                # assert f_style is not None

                Lo.delay(1_000)
                doc.close()
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

The ``style_text_attributes()`` method is called to set the various boolean options in the ``Text Attributes`` section of the Chart Data Labels dialog as seen in :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code

        ds = chart_doc.get_data_series()[0]
        ds.style_text_attributes(
            show_category_name=True,
            show_legend_symbol=True,
            show_series_name=True,
            auto_text_wrap=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`ffd2621d-fb71-4a00-ad8a-5d0760ed11bf_1` and :numref:`7852b8b7-054a-484c-823c-3512b700132b_1`.

.. cssclass:: screen_shot

    .. _ffd2621d-fb71-4a00-ad8a-5d0760ed11bf_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ffd2621d-fb71-4a00-ad8a-5d0760ed11bf
        :alt: Chart with formatting applied to data series
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data series

.. cssclass:: screen_shot

    .. _7852b8b7-054a-484c-823c-3512b700132b_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/7852b8b7-054a-484c-823c-3512b700132b
        :alt: Chart Format Number Dialog
        :figclass: align-center
        :width: 450px

        Chart Format Number Dialog

Style Data Point
~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[2]

        dp.style_text_attributes(
            show_category_name=True,
            show_legend_symbol=True,
            show_series_name=True,
            auto_text_wrap=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`959761b7-4336-4712-8e86-a48897606925_1`.

.. cssclass:: screen_shot

    .. _959761b7-4336-4712-8e86-a48897606925_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/959761b7-4336-4712-8e86-a48897606925
        :alt: Chart with Text Attributes applied to data point
        :figclass: align-center
        :width: 450px

        Chart with Text Attributes applied to data point

Number Format
"""""""""""""

The ``style_numbers_numbers()`` method is used to set the number format of the data labels.
This method is used to set the values seen in :numref:`ca21f3f1-e1b1-4bab-bb36-f52c966e00af`.

The ``NumberFormatIndexEnum`` enum contains the values in |num_fmt_index|_ for easy lookup.

To ensure that the effects of ``style_numbers_numbers()`` are visible the ``style_text_attributes()`` method is called to
turn on ``Value as Number`` of the dialog seen in :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
        # ... other code

        ds.style_text_attributes(show_number=True)
        ds.style_numbers_numbers(
            source_format=False,
            num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`3d1f582b-558d-4da5-8996-bebb6b6781d0_1` and :numref:`ca21f3f1-e1b1-4bab-bb36-f52c966e00af_1`.

.. cssclass:: screen_shot

    .. _3d1f582b-558d-4da5-8996-bebb6b6781d0_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3d1f582b-558d-4da5-8996-bebb6b6781d0
        :alt: Chart with Text Attributes applied to data series
        :figclass: align-center
        :width: 450px

        Chart with Text Attributes applied to data series

.. cssclass:: screen_shot

    .. _ca21f3f1-e1b1-4bab-bb36-f52c966e00af_1:

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
        ds = chart_doc.get_data_series()[0]
        dp = ds[1]
        dp.style_text_attributes(show_number=True)
        dp.style_numbers_numbers(
            source_format=False,
            num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`157ea466-4511-4f84-90e9-52b76390c1fb_1`.

.. cssclass:: screen_shot

    .. _157ea466-4511-4f84-90e9-52b76390c1fb_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/157ea466-4511-4f84-90e9-52b76390c1fb
        :alt: Chart with Text Attributes applied to data point
        :figclass: align-center
        :width: 450px

        Chart with Text Attributes applied to data point

Percentage Format
"""""""""""""""""

The ``style_numbers_percent()`` method is called to set the number format of the data labels.
This class is used to set the values seen in :numref:`ca21f3f1-e1b1-4bab-bb36-f52c966e00af`.

The ``NumberFormatIndexEnum`` enum contains the values in |num_fmt_index|_ for easy lookup.

To ensure that the effects of ``style_numbers_percent()`` are visible the ``style_text_attributes()``
method is called to turn on ``Value as Percentage`` of the dialog seen in :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
        # ... other code

        ds = chart_doc.get_data_series()[0]
        ds.style_text_attributes(show_number_in_percent=True)
        ds.style_numbers_percent(
            source_format=False,
            num_format_index=NumberFormatIndexEnum.PERCENT_DEC2,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`d8b1329b-d94e-457d-91d1-87d5f14aefa2_1` and :numref:`45c0d0a1-4c9e-4b84-ad9b-c92bb4a2658e_1`.

.. cssclass:: screen_shot

    .. _d8b1329b-d94e-457d-91d1-87d5f14aefa2_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d8b1329b-d94e-457d-91d1-87d5f14aefa2
        :alt: Chart with formatting applied to data series
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data series

.. cssclass:: screen_shot

    .. _45c0d0a1-4c9e-4b84-ad9b-c92bb4a2658e_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/45c0d0a1-4c9e-4b84-ad9b-c92bb4a2658e
        :alt: Chart Format Number Dialog
        :figclass: align-center
        :width: 450px

        Chart Format Number Dialog

Style Data Point
~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[3]
        dp.style_text_attributes(show_number_in_percent=True)
        dp.style_numbers_percent(
            source_format=False,
            num_format_index=NumberFormatIndexEnum.PERCENT_DEC2,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`cc247b93-54e7-4f51-a5c7-c80c759eaad8_1`.

.. cssclass:: screen_shot

    .. _cc247b93-54e7-4f51-a5c7-c80c759eaad8_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/cc247b93-54e7-4f51-a5c7-c80c759eaad8
        :alt: Chart with formatting applied to data point
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data point

Attribute Options
-----------------

The ``style_attribute_options()`` method is called to set the Options data labels.
This class is used to set the values seen in the ``Attribute Options`` section of :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

The :py:class:`~ooodev.format.chart2.direct.series.data_labels.data_labels.PlacementKind` enum is used to look up the placement.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.attrib_options import PlacementKind
        # ... other code

        ds = chart_doc.get_data_series()[0]
        ds.style_attribute_options(placement=PlacementKind.INSIDE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`115e2eaa-876c-4048-b30a-06e5be91b240_1` and :numref:`6b9458d9-b457-4de2-aa54-7c44a711e2a2_1`.

.. cssclass:: screen_shot

    .. _115e2eaa-876c-4048-b30a-06e5be91b240_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/115e2eaa-876c-4048-b30a-06e5be91b240
        :alt: Chart with formatting applied to data series
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data series

.. cssclass:: screen_shot

    .. _6b9458d9-b457-4de2-aa54-7c44a711e2a2_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6b9458d9-b457-4de2-aa54-7c44a711e2a2
        :alt: Chart Format Number Dialog
        :figclass: align-center
        :width: 450px

        Chart Format Number Dialog

Style Data Point
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.attrib_options import PlacementKind
        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[-1]  # get the last data point
        dp.style_attribute_options(placement=PlacementKind.INSIDE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`4968c491-5e45-449e-800f-01549bc009bd_1`.

.. cssclass:: screen_shot

    .. _4968c491-5e45-449e-800f-01549bc009bd_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4968c491-5e45-449e-800f-01549bc009bd
        :alt: Chart with formatting applied to data point
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data point

Rotate Text
-----------

The ``style_orientation()`` method is called to set the rotation of data labels.
This class is used to set the values seen in the ``Rotate Text`` section of :numref:`0c7d3398-34f5-4da3-81f2-79e134fab44d`.

The :py:class:`~ooodev.format.inner.direct.chart2.title.alignment.direction.DirectionModeKind` enum is used to look up the text direction.

Before formatting the chart is seen in :numref:`236874763-f2b763db-c294-4496-971e-d4982e6d7b68`.

Style Data Series
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind
        # ... other code
        ds = chart_doc.get_data_series()[0]
        ds.style_orientation(angle=60, mode=DirectionModeKind.LR_TB, leaders=True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`d57bc634-0f1e-4acc-9d02-848809635021_1` and :numref:`91cac9f6-9dbb-4017-a682-cd7a977c208e_1`.

.. cssclass:: screen_shot

    .. _d57bc634-0f1e-4acc-9d02-848809635021_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d57bc634-0f1e-4acc-9d02-848809635021
        :alt: Chart with formatting applied to data series
        :figclass: align-center
        :width: 450px

        Chart with formatting applied to data series

.. cssclass:: screen_shot

    .. _91cac9f6-9dbb-4017-a682-cd7a977c208e_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/91cac9f6-9dbb-4017-a682-cd7a977c208e
        :alt: Chart Format Number Dialog
        :figclass: align-center
        :width: 450px

        Chart Format Number Dialog

Style Data Point
""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind
        # ... other code
        ds = chart_doc.get_data_series()[0]
        dp = ds[2]
        dp.style_orientation(angle=60, mode=DirectionModeKind.LR_TB, leaders=True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`35ff95c1-f3b3-48d6-848f-8c2935faa9b3_1`

.. cssclass:: screen_shot

    .. _35ff95c1-f3b3-48d6-848f-8c2935faa9b3_1:

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
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_general`
        - |num_fmt|_
        - |num_fmt_index|_
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`

.. |num_fmt| replace:: API NumberFormat
.. _num_fmt: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html

.. |num_fmt_index| replace:: API NumberFormatIndex
.. _num_fmt_index: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html
