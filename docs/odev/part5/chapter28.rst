.. _ch28:

***************************************
Chapter 28. Functions and Data Analysis
***************************************

.. topic:: Overview

    Charting Elements; Chart Creation: TableChart_, |chart2_srv|_, linking template, diagram, and data source; Modifying Chart Elements: diagram, coordinate system, chart type, data series

At over 1,600 pages the OpenOffice Developer's Guide isn't a quick read, but you might expect it to cover all the major parts of the Office API.
That's mostly true, except for one omission - there's no mention of the ``chart2`` charting module.
It's absent from the guide, the online examples, and only makes a brief appearance in the Wiki, at `<https://wiki.openoffice.org/wiki/Chart2>`__.

That's not to say that chart creation isn't explained in the guide; chapter 10 is all about that topic, but using the older charting module, called chart (as you might guess).
One source of confusion is that both modules (``chart`` and ``chart2``) have a similar top-level interface to Calc via XTableChart_ and ``XChartDocument``,
but they rapidly diverge as you progress further into the APIs.
A sure way to make code crash is to mix services and interfaces from the two modules.

Since newer is obviously better, the question arises as to why the Developer's Guide skips ``chart2``?
The reason seems to be historical - the guide was written for OpenOffice version ``3.1``, which dates from the middle of 2009.
The ``chart2`` module was released two years before, in September 2007, for version ``2.3``.
That release came with dire warnings in the Wiki about the API being unstable and subject to change.
I'm sure those warnings were valid back in 2007, but ``chart2`` underwent a lot of development over the next three years before LibreOffice was forked off in September 2010.
After that the pace of change slowed, mainly because the module was stable.
For example, Calc's charting wizard is implemented using ``chart2``.

Since the developer's guide hasn't been updated in since 2011, the ``chart2`` module hasn't received much notice.
|odev| is rectifying that by concentrating solely on `chart2` programming; The old chart API is not used here, although |odev| has a :py:class:`~.chart.Chart` class.

The primary source of online information about ``chart2`` is its `API documentation <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1chart2.html>`__.

One problem with searching the Web for examples is that programs using the ``chart2`` and ``chart`` modules look similar.
The crucial difference for Python is that most of the ``chart2`` services and interfaces are inside the ``com.sun.star.chart2`` package whereas
the older ``chart`` services and interfaces are inside ``com.sun.star.chart``.

Another way to distinguish between examples, especially for programs written in Basic, is to look at the names of the chart services.
The old chart names end with the word ``Diagram`` (:abbreviation:`i.e.` ``BarDiagram``, ``DonutDiagram``, ``LineDiagram``) whereas the ``chart2`` names either end with
the word ``ChartType`` (:abbreviation:`i.e.` ``BarChartType``, ``PieChartType``, ``LineChartType``), or with no special word (:abbreviation:`i.e.` ``Bar``, ``Donut``, ``Line``).

A good way to get a feel for ``chart2's`` functionality is to look at |ug_ch03|_ of the |ug|_, available from `<https://libreoffice.org/get-help/documentation>`__.
It describes the charting wizard, which acts as a very nice GUI for most of the ``chart2`` API.
The chapter also introduces a lot of charting terminology (:abbreviation:`i.e.` chart types, data ranges, data series) used in the API.

.. _ch28_charting_elements:

28.1 Charting Elements
======================

Different chart types share many common elements, as illustrated in :numref:`ch28fig_typical_chart_el`.

..
    figure 1

.. cssclass:: diagram invert

    .. _ch28fig_typical_chart_el:
    .. figure:: https://user-images.githubusercontent.com/4193389/206014490-1bc18216-9993-4b81-a9a0-69f5656dd7c4.png
        :alt: Typical Chart Elements.
        :figclass: align-center

        :Typical Chart Elements.

Most of the labeled components in :numref:`ch28fig_typical_chart_el` are automatically included when a chart template is instantiated;
the programmer typically only has to supply text and a few property settings, such as the wall color and font size.

There are ten main chart types, which are listed when the "Chart Wizard" is started via Calc's Insert, Chart menu item.
:numref:`ch28fig_ten_chart_types` shows the possibilities.

..
    figure 2

.. cssclass:: screen_shot invert

    .. _ch28fig_ten_chart_types:
    .. figure:: https://user-images.githubusercontent.com/4193389/206015606-6c894677-24da-4023-91c8-72e9f17dbb82.png
        :alt: Ten Chart Types
        :figclass: align-center

        :Ten Chart Types.

Most of the types offer variants, which are displayed as icons to the right of the dialog window.
When you position the mouse over an icon, the name of the variant appears briefly in a tooltip, as in :numref:`ch28fig_column_chart_tooltip_name`.

..
    figure 3

.. cssclass:: screen_shot invert

    .. _ch28fig_column_chart_tooltip_name:
    .. figure:: https://user-images.githubusercontent.com/4193389/206016079-01926c4e-2ee0-450a-a22a-6f8dcd7c05a2.png
        :alt: A Column Chart Icon with its Tooltip Name
        :figclass: align-center

        :A Column Chart Icon with its Tooltip Name.

When the :spelling:word:`checkboxes`, buttons, or combo boxes are selected in the dialog, the icons change to reflect changes in the variants.

The three most common variants are ``3D``, ``Stacked`` and ``Percent``. ``Stacked`` is utilized when the chart displays multiple data sequences stacked on top of each other.
``Percent`` is combined with ``Stacked`` to stack the sequences in terms of their percentage contribution to the total.
A lengthy discussion about chart variants can be found in |ug_ch03|_ of the |ug|_, in the section `Gallery of Chart Types <https://books.libreoffice.org/en/CG74/CG7403-ChartsAndGraphs.html#toc116>`__.

In the ``chart2`` API, the variants are accessed via template names, which are listed in :numref:`ch28tblchart_types_and_template_names`.

.. _ch28tblchart_types_and_template_names:

.. table:: Chart Types and Template Names
    :name: chart_types_and_template_names
    :align: center
    :class: ul-list

    +------------------+----------+----------------------------------+
    | Chart            | Types    | Template Names                   |
    +==================+==========+==================================+
    | Column           | Stacked  | - Column                         |
    |                  |          | - StackedColumn                  |
    +------------------+----------+----------------------------------+
    |                  | Percent  | - ThreeDColumnDeep               |
    |                  |          | - ThreeDColumnFlat               |
    +------------------+----------+----------------------------------+
    |                  | 3D       | - StackedThreeDColumnFlat        |
    |                  |          | - PercentStackedThreeDColumnFlat |
    +------------------+----------+----------------------------------+
    | Bar              | Stacked  | - Bar                            |
    |                  |          | - StackedBar                     |
    |                  |          | - PercentStackedBar              |
    +------------------+----------+----------------------------------+
    |                  | Percent  | - ThreeDBarDeep                  |
    |                  |          | - ThreeDBarFlat                  |
    +------------------+----------+----------------------------------+
    |                  | 3D       | - StackedThreeDBarFlat           |
    |                  |          | - PercentStackedThreeDBarFlat    |
    +------------------+----------+----------------------------------+
    | Pie              | Donut    | - Pie                            |
    |                  |          | - Donut                          |
    +------------------+----------+----------------------------------+
    |                  | Explode  | - PieAllExploded                 |
    |                  |          | - DonutAllExploded               |
    +------------------+----------+----------------------------------+
    |                  | 3D       | - ThreeDPie                      |
    |                  |          | - ThreeDPieAllExploded           |
    |                  |          | - ThreeDDonut                    |
    |                  |          | - ThreeDDonutAllExploded         |
    +------------------+----------+----------------------------------+
    | Area             | Stacked  | - Area                           |
    |                  |          | - StackedArea                    |
    |                  |          | - PercentStackedArea             |
    +------------------+----------+----------------------------------+
    | Area             | Stacked  | - Area                           |
    |                  |          | - StackedArea                    |
    |                  |          | - PercentStackedArea             |
    +------------------+----------+----------------------------------+
    |                  | Percent  | - ThreeDArea                     |
    |                  |          | - StackedThreeDArea              |
    +------------------+----------+----------------------------------+
    |                  | 3D       | - PercentStackedThreeDArea       |
    +------------------+----------+----------------------------------+
    | Line             | Symbol   | - Line                           |
    |                  |          | - Symbol                         |
    |                  |          | - LineSymbol                     |
    +------------------+----------+----------------------------------+
    |                  | Stacked  | - StackedLine                    |
    |                  |          | - StackedSymbol                  |
    |                  |          | - StackedLineSymbol              |
    +------------------+----------+----------------------------------+
    |                  | Percent  | - PercentStackedLine             |
    |                  |          | - PercentStackedSymbol           |
    +------------------+----------+----------------------------------+
    |                  | 3D       | - PercentStackedLineSymbol       |
    |                  |          | - ThreeDLine                     |
    |                  |          | - ThreeDLineDeep                 |
    |                  |          | - StackedThreeDLine              |
    |                  |          | - PercentStackedThreeDLine       |
    +------------------+----------+----------------------------------+
    | XY               | Line     | - ScatterSymbol                  |
    |                  |          | - ScatterLine                    |
    |                  |          | - ScatterLineSymbol              |
    +------------------+----------+----------------------------------+
    | (Scatter)        | 3D       | - ThreeDScatter                  |
    +------------------+----------+----------------------------------+
    | Bubble           |          | - Bubble                         |
    +------------------+----------+----------------------------------+
    | Net              | Line     | - Net                            |
    |                  |          | - NetLine                        |
    |                  |          | - NetSymbol                      |
    |                  |          | - FilledNet                      |
    +------------------+----------+----------------------------------+
    |                  | Symbol   | - StackedNet                     |
    |                  |          | - StackedNetLine                 |
    +------------------+----------+----------------------------------+
    |                  | Filled   | - StackedNetSymbol               |
    |                  |          | - StackedFilledNet               |
    +------------------+----------+----------------------------------+
    |                  | Stacked  | - PercentStackedNet              |
    |                  |          | - PercentStackedNetLine          |
    |                  |          | - PercentStackedNetSymbol        |
    +------------------+----------+----------------------------------+
    |                  | Percent  | - PercentStackedFilledNet        |
    +------------------+----------+----------------------------------+
    | Stock            | Open     | - StockLowHighClose              |
    +------------------+----------+----------------------------------+
    |                  | Volume   | - StockOpenLowHighClose          |
    |                  |          | - StockVolumeLowHighClose        |
    |                  |          | - StockVolumeOpenLowHighClose    |
    +------------------+----------+----------------------------------+
    | Column and Line  | Stacked  | - ColumnWithLine                 |
    |                  |          | - StackedColumnWithLine          |
    +------------------+----------+----------------------------------+

The template names are closely related to the tooltip names in Calc's chart wizard.
For example, the tooltip name in :numref:`ch28fig_column_chart_tooltip_name` corresponds to the ``PercentStackedColumn`` template.

It's also possible to create a chart using a chart type name, which are listed in :numref:`ch28tbl_chart_type_names`.

.. _ch28tbl_chart_type_names:

.. table:: Chart Type Names.
    :name: chart_type_names
    :align: center

    ======== ====================================
     Chart    Chart Type Names                   
    ======== ====================================
     Column   ColumnChartType
     Bar      BarChartType
     Pie      PieChartType
     Area     AreaChartType
     Line     LineChartType
     XY       (Scatter) ScatterChartType
     Bubble   BubbleChartType
     Net      NetChartType, FilledNetChartType
     Stock    CandleStickChartType
    ======== ====================================

|odev| has :py:class:`~.kind.chart2_types.ChartTypes` class for looking up chart names to make it a bit easier for a developer.
:py:class:`~.kind.chart2_types.ChartTypes` is has a sub-class for each chart type shown in :numref:`ch28tbl_chart_type_names`.
Each sub-class has a ``NAMED`` field which contain the name in column ``2`` of :numref:`ch28tbl_chart_type_names`.
Also each sub-class has one or more fileds that start with ``TEMPLATE_`` such as ``TEMPLATE_3D`` or ``TEMPLATE_PERCENT``.
``TEMPLATE_`` fields point to the possible chart template names listed in column ``3`` of :numref:`ch28tblchart_types_and_template_names`.

For Example ``diagram_name`` of :py:meth:`.Chart2.insert_chart` can be passeed ``ChartTypes.Pie.TEMPLATE_DONUT.DONUT``.

.. tabs::

    .. code-tab:: python

        range_addr = Calc.get_address(sheet=sheet, range_name="A44:C50")
        chart_doc = Chart2.insert_chart(
            sheet=sheet,
            cells_range=range_addr,
            cell_name="D43",
            width=15,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_DONUT.DONUT,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Note that a stock chart graph is drawn using a ``CandleStickChartType``, and that there's no type name for a column and line chart because it's implemented as a combination of ``ColumnChartType`` and ``BarChartType``.

The ``chart2`` module is quite complex, so |odev| hides a lot of details inside methods in :py:class:`~.chart2.Chart2` class. It simplifies four kinds of operation:

1. The creation of a new chart in a spreadsheet document, based on a template name.
2. The accessing and modification of elements inside a chart, such as the title, legend, axes, and colors.
3. The addition of extra data to a chart, such as error bars or a second graph.
4. The embedding of a chart in a document other than a spreadsheet, namely in a text document or slide presentation.

Operations no. 1 (chart creation) and no. 2 (element modification) are used by all my examples, so the rest of this chapter will give an overview of how the corresponding :py:class:`~.chart2.Chart2` methods work.

Programming details specific to particular charts will be discussed in subsequent chapters:

.. todo::

    | Chapter 28, Add link to chapters 30
    | Chapter 28, Add link to chapters 31
    | Chapter 28, Add link to chapters 32

.. cssclass:: ul-list

    - column: chapter 29;
    - bar, pie, area, line: chapter 30;
    - XY (scatter): chapter 31;
    - bubble, net, stock: chapter 32.

.. _ch28_chart_creation:

28.2 Chart Creation
===================

Chart creation can be divided into three steps:

1. A TableChart_ service is created inside the spreadsheet.
2. The |chart2_srv|_ service is accessed inside the TableChart_.
3. The |chart2_srv|_ is initialized by linking together a chart template, diagram, and data source.

The details are explained in the following sub-sections.

.. _ch28_creating_tbl_chart:

28.2.1 Creating a Table Chart
-----------------------------

Work in progress ...


.. |chart2_srv| replace:: ChartDocument
.. _chart2_srv: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1ChartDocument.html

.. |XChartDocument2| replace:: XChartDocument
.. _XChartDocument2: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XChartDocument.html

.. |ug| replace:: Calc User Guide
.. _ug: https://books.libreoffice.org/en/CG74/CG74.html

.. |ug_ch03| replace:: chapter 3
.. _ug_ch03: https://books.libreoffice.org/en/CG74/CG7403-ChartsAndGraphs.html

.. _TableChart: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableChart.html
.. _XTableChart: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XTableChart.html


.. spelling:word-list::
    Donut
