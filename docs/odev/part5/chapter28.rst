.. _ch28:

***************************************
Chapter 28. Functions and Data Analysis
***************************************

.. topic:: Overview

    Charting Elements; Chart Creation: TableChart_, |ChartDocument2|_, linking template, diagram, and data source; Modifying Chart Elements: diagram, coordinate system, chart type, data series

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

..
    Table 1

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

..
    Table 2

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
Also each sub-class has one or more fields that start with ``TEMPLATE_`` such as ``TEMPLATE_3D`` or ``TEMPLATE_PERCENT``.
``TEMPLATE_`` fields point to the possible chart template names listed in column ``3`` of :numref:`ch28tblchart_types_and_template_names`.

For Example ``diagram_name`` of :py:meth:`.Chart2.insert_chart` can be passed ``ChartTypes.Pie.TEMPLATE_DONUT.DONUT``.

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
2. The |ChartDocument2|_ service is accessed inside the TableChart_.
3. The |ChartDocument2|_ is initialized by linking together a chart template, diagram, and data source.

The details are explained in the following sub-sections.

.. _ch28_creating_tbl_chart:

28.2.1 Creating a Table Chart
-----------------------------

``XTableCharts.addNewByName()`` adds a new TableChart_ to the TableCharts_ collection in a spreadsheet.
This is shown graphically in :numref:`ch28fig_new_tablechart`, and is implemented by :py:meth:`.Chart2.add_table_chart`.

..
    figure 4

.. cssclass:: diagram invert

    .. _ch28fig_new_tablechart:
    .. figure:: https://user-images.githubusercontent.com/4193389/206303477-20539205-2885-4957-9b4e-854990cae5f9.png
        :alt: Creating a new TableChart Service
        :figclass: align-center

        :Creating a new TableChart_ Service.

:py:meth:`.Chart2.add_table_chart` is defined as:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def add_table_chart(
            sheet: XSpreadsheet, chart_name: str, cells_range: CellRangeAddress,
            cell_name: str, width: int, height: int
        ) -> None:
            try:
                charts_supp = Lo.qi(XTableChartsSupplier, sheet, True)
                tbl_charts = charts_supp.getCharts()

                pos = Calc.get_cell_pos(sheet, cell_name)
                rect = Rectangle(X=pos.X, Y=pos.Y, Width=width * 1_000, Height=height * 1_000)
                addrs = (cells_range,)

                tbl_charts.addNewByName(chart_name, rect, addrs, True, True)
            except Exception as e:
                raise ChartError("Error adding table chart") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The arguments passed to :py:meth:`.Chart2.add_table_chart` include the new chart's name, the cell range used as a data source, and the chart's position and dimensions when drawn in the Calc window.

The position is a cell name (:abbreviation:`i.e.` ``A1``), which becomes the location of the top-left corner of the chart in the Calc window.
The name is converted into a position by calling :py:meth:`.Calc.get_cell_pos`.
The size of the chart is supplied as millimeter width and height arguments and converted into a Rectangle in ``1/100mm`` units.

The methods assume that the data range has a specific format, which is illustrated by :numref:`ch28fig_cell_rng_data_fmt`.

..
    figure 5

.. cssclass:: screen_shot invert

    .. _ch28fig_cell_rng_data_fmt:
    .. figure:: https://user-images.githubusercontent.com/4193389/206309482-21489f85-a986-4a39-854a-c10784d44f8a.png
        :alt: Cell Range Data Format
        :figclass: align-center

        :Cell Range Data Format.

The data is organized into columns, the first for the ``x-axis`` categories, and the others for the ``y-axis`` data displayed as graphs.
The first row of the data range contains labels for the ``x-axis`` and the graphs.

For example, the data range in :numref:`ch28fig_cell_rng_data_fmt` is drawn as a Column chart in :numref:`ch28fig_colum_chart_via_fig5`.

..
    figure 6

.. cssclass:: screen_shot

    .. _ch28fig_colum_chart_via_fig5:
    .. figure:: https://user-images.githubusercontent.com/4193389/206310637-43a45c2a-ab86-483e-b837-e4185db1711e.png
        :alt: A Column Chart Using the Data in previous figure.
        :figclass: align-center

        :A Column Chart Using the Data in :numref:`ch28fig_cell_rng_data_fmt`.

The assumption that the first data column are ``x-axis`` categories doesn't apply to scatter and bubble charts which use numerical ``x-axis`` values.
There are examples of those in later chapters.

The data format assumptions are used in the call to ``XTableCharts.addNewByName()`` in :py:meth:`.Chart2.add_table_chart` by setting its last two arguments to ``True``.
This specifies that the top row and left column will be used as categories and/or labels.
More specific format information will be supplied later.

.. _ch28_accessing_chart_doc:

28.2.2 Accessing the Chart Document
-----------------------------------

Although :py:meth:`.Chart2.add_table_chart` adds a table chart to the spreadsheet, it doesn't return a reference to the new chart document.
That's obtained by calling :py:meth:`.Chart2.get_chart_doc`:

.. tabs::

    .. code-tab:: python

        Chart2.add_table_chart(
            sheet=sheet,
            chart_name=chart_name,
            cells_range=cells_range,
            cell_name=cell_name,
            width=width,
            height=height
        )
        chartDoc = Chart2.get_chart_doc(sheet=sheet, chart_name=chartName) # XChartDocument

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.get_chart_doc` accesses the spreadsheet's collection of TableCharts_, searching for the one with the given name.
The matching TableChart_ service is treated as an XEmbeddedObjectSupplier_ interface, which lets its embedded chart document be referenced.
These steps are illustrated by :numref:`ch28fig_acc_chart_doc`.

..
    figure 7

.. cssclass:: diagram invert

    .. _ch28fig_acc_chart_doc:
    .. figure:: https://user-images.githubusercontent.com/4193389/206313332-a1cd22cc-4a2a-49e3-bb04-44777ca59837.png
        :alt: Accessing a Chart Document.
        :figclass: align-center

        :Accessing a Chart Document.

:py:meth:`.Chart2.get_chart_doc` implements :numref:`ch28fig_acc_chart_doc`, using :py:meth:`.Chart2.get_table_chart` to access the named table chart:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def get_chart_doc(cls, sheet: XSpreadsheet, chart_name: str) -> XChartDocument:
            try:
                tbl_chart = cls.get_table_chart(sheet, chart_name)
                eos = Lo.qi(XEmbeddedObjectSupplier, tbl_chart, True)
                return Lo.qi(XChartDocument, eos.getEmbeddedObject(), True)
            except ChartError:
                raise
            except Exception as e:
                raise ChartError(f'Error getting chart document for chart "{chart_name}"') from e

        @staticmethod
        def get_table_chart(sheet: XSpreadsheet, chart_name: str) -> XTableChart:
            try:
                charts_supp = Lo.qi(XTableChartsSupplier, sheet, True)
                tbl_charts = charts_supp.getCharts()
                tc_access = Lo.qi(XNameAccess, tbl_charts, True)
                tbl_chart = Lo.qi(XTableChart, tc_access.getByName(chart_name))
                return tbl_chart
            except Exception as e:
                raise ChartError(f'Error getting table chart for chart "{chart_name}"') from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch28_initalizing_chart_doc:

28.2.3 Initializing the Chart Document
--------------------------------------

The chart document is initialized by linking three components: the chart template, the chart's diagram, and a data source, as illustrated by :numref:`ch28fig_initalizing_chart_doc`.

..
    figure 8

.. cssclass:: diagram invert

    .. _ch28fig_initalizing_chart_doc:
    .. figure:: https://user-images.githubusercontent.com/4193389/206314319-89b70bdd-33d3-461b-b609-b307ffa78616.png
        :alt: Initializing a Chart Document
        :figclass: align-center

        :Initializing a Chart Document.

The initialization steps in :numref:`ch28fig_initalizing_chart_doc`, and the earlier calls to :py:meth:`.Chart2.add_table_chart` and :py:meth:`.Chart2.get_chart_doc` are carried out by :py:meth:`.Chart2.insert_chart`.
A typical call to ``insert_chart()`` would be:

.. tabs::

    .. code-tab:: python

        range_addr = Calc.get_address(sheet=sheet, range_name="E15:G21") # CellRangeAddress
        chart_doc =  Chart2.insert_chart(
            sheet=sheet, 
            cells_range=range_addr,
            cell_name="A22",
            width=20,
            height=11,
            diagram_name=ChartTypes.Column.TEMPLATE_STACKED.COLUMN # or "Column"
        ) # XChartDocument

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The first line converts ``E15:G21`` into a data range (this corresponds to the cells shown in :numref:`ch28fig_cell_rng_data_fmt`), which is passed to :py:meth:`.Chart2.insert_chart`.
The ``A22`` string and the ``20x11 mm`` dimensions specify the position and size of the chart, and the last argument (``Column``)
is the desired chart template (see :numref:`ch28tblchart_types_and_template_names`, see :py:class:`~.kind.chart2_types.ChartTypes`).
The result is the column chart shown in :numref:`ch28fig_colum_chart_via_fig5`.

:py:meth:`.Chart2.insert_chart` is:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def insert_chart(
            cls,
            sheet: XSpreadsheet,
            cells_range: CellRangeAddress,
            cell_name: str,
            width: int,
            height: int,
            diagram_name: ChartTemplateBase | str,
            color_bg: Color = mColor.CommonColor.PALE_BLUE,
            color_wall: Color = mColor.CommonColor.LIGHT_BLUE,
        ) -> XChartDocument:
            try:
                # type check that diagram_name is ChartTemplateBase | str
                Info.is_type_enum_multi(
                    alt_type="str", enum_type=ChartTemplateBase,
                    enum_val=diagram_name, arg_name="diagram_name"
                )
                chart_name = Chart2._CHART_NAME + str(int(random() * 10_000))
                cls.add_table_chart(
                    sheet=sheet,
                    chart_name=chart_name,
                    cells_range=cells_range,
                    cell_name=cell_name,
                    width=width,
                    height=height,
                )
                chart_doc = cls.get_chart_doc(sheet, chart_name)

                # assign chart template to the chart's diagram
                diagram = chart_doc.getFirstDiagram()
                ct_template = cls.set_template(
                    chart_doc=chart_doc, diagram=diagram, diagram_name=diagram_name
                )

                has_cats = cls.has_categories(diagram_name)

                dp = chart_doc.getDataProvider()

                ps = Props.make_props(
                    CellRangeRepresentation=Calc.get_range_str(cells_range, sheet),
                    DataRowSource=ChartDataRowSource.COLUMNS,
                    FirstCellAsLabel=True,
                    HasCategories=has_cats,
                )
                ds = dp.createDataSource(ps)

                # add data source to chart template
                args = Props.make_props(HasCategories=has_cats)
                ct_template.changeDiagramData(diagram, ds, args)

                # apply style settings to chart doc
                # background and wall colors
                cls.set_background_colors(chart_doc, color_bg, color_wall)

                if has_cats:
                    cls.set_data_point_labels(chart_doc, DataPointLabelTypeKind.NUMBER)

                return chart_doc
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error inserting chart") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`~.Chart2.insert_chart` creates a new chart document by calling :py:meth:`~.Chart2.add_table_chart` and :py:meth:`~.Chart2.get_chart_doc`,
and then proceeds to link the chart template, diagram, and data source.

Get the Diagram
^^^^^^^^^^^^^^^

The chart diagram is the easiest to obtain, since it's directly accessible via the |XChartDocument2|_ reference:

.. tabs::

    .. code-tab:: python

        # part of Chart2.insert_chart()...
        diagram = chart_doc.getFirstDiagram() # XDiagram

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Creating a Template
^^^^^^^^^^^^^^^^^^^

Creating a chart template is a few more steps. requiring the creation of a XChartTypeManager_ interface inside :py:meth:`.Chart2.set_template`:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def set_template(
            chart_doc: XChartDocument, diagram: XDiagram, diagram_name: ChartTemplateBase | str
        ) -> XChartTypeTemplate:

            # ensure diagram_name is ChartTemplateBase | str
            Info.is_type_enum_multi(
                alt_type="str", enum_type=ChartTemplateBase, enum_val=diagram_name, arg_name="diagram_name"
            )

            try:
                ct_man = chart_doc.getChartTypeManager()
                msf = Lo.qi(XMultiServiceFactory, ct_man, True)
                template_nm = f"com.sun.star.chart2.template.{diagram_name}"
                ct_template = Lo.qi(XChartTypeTemplate, msf.createInstance(template_nm))
                if ct_template is None:
                    Lo.print(
                        f'Could not create chart template "{diagram_name}"; using a column chart instead'
                    )
                    ct_template = Lo.qi(
                        XChartTypeTemplate, msf.createInstance("com.sun.star.chart2.template.Column"), True
                    )

                ct_template.changeDiagram(diagram)
                return ct_template
            except Exception as e:
                raise ChartError("Error setting chart template") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The ``diagram_name`` value is one of the template names shown in :numref:`ch28tblchart_types_and_template_names` (:abbreviation:`i.e.` ``Column``).
The string ``com.sun.star.chart2.template.`` is added to the front to create a fully qualified service name, which is then instantiated.
If the instance creation fails, then the function falls back to creating an instance of the ``Column`` template.
:py:meth:`~.Chart2.set_template` ends by calling ``XChartTypeTemplate.changeDiagram()`` which links the template to the chart's diagram.

Get the Data Source
^^^^^^^^^^^^^^^^^^^

Back in :py:meth:`.Chart2.insert_chart`, the right-most branch of :numref:`ch28fig_initalizing_chart_doc` involves the creation of an XDataProvider_ instance:

.. tabs::

    .. code-tab:: python

        # part of Chart2.insert_chart()...
        dp = chart_doc.getDataProvider() # XDataProvider

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This data provider converts the chart's data range into an XDataSource_:

.. tabs::

    .. code-tab:: python

        # part of Chart2.insert_chart()...
        has_cats = cls.has_categories(diagram_name)

        ps = Props.make_props(
            CellRangeRepresentation=Calc.get_range_str(cells_range, sheet),
            DataRowSource=ChartDataRowSource.COLUMNS,
            FirstCellAsLabel=True,
            HasCategories=has_cats,
        )
        ds = dp.createDataSource(ps) # XDataSource

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The properties passed to ``XDataProvider.createDataSource()`` specify more details about the format of the data in
:numref:`ch28fig_cell_rng_data_fmt` - the data for each graph is organized into columns with the first cell being the label for the graph.
The ``HasCategories`` property is set to true when the first column of the data is to be used as ``x-axis`` categories.

These properties passed to ``createDataSource()`` are described in the documentation for the TabularDataProviderArguments_ service.

The ``has_cats`` boolean is set by examining the diagram name: if it's an XY scatter chart or bubble chart then
the first column of data will not be used as ``x-axis`` categories, so the boolean is set to ``False``:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def has_categories(diagram_name: ChartTemplateBase | str) -> bool:
            # Ensure diagram_name ChartTemplateBase | str
            Info.is_type_enum_multi(
                alt_type="str", enum_type=ChartTemplateBase, enum_val=diagram_name, arg_name="diagram_name"
            )

            dn = str(diagram_name).lower()
            non_cats = ("scatter", "bubble")
            for non_cat in non_cats:
                if non_cat in dn:
                    return False
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Linking the template, diagram, and data source
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Now the data source can populate the diagram using the specified chart template format:

.. tabs::

    .. code-tab:: python

        # part of Chart2.insert_chart()...
        # add data source to chart template
        args = Props.make_props(HasCategories=has_cats)
        ct_template.changeDiagramData(diagram, ds, args)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

At this point the chart will be drawn in the Calc application window, and :py:meth:`.Chart2.insert_chart` could return.
Instead my code modifies the appearance of the chart in two ways:

.. tabs::

    .. code-tab:: python

        # part of Chart2.insert_chart()...
        # apply style settings to chart doc
        # background and wall colors
        cls.set_background_colors(chart_doc, color_bg, color_wall)

        if has_cats:  # charts using x-axis categories
            cls.set_data_point_labels(chart_doc, DataPointLabelTypeKind.NUMBER)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.set_background_colors` changes the background and wall colors of the chart (see :numref:`ch28fig_colum_chart_via_fig5`).
:py:meth:`.Chart2.set_data_point_labels` switches on the displaying of the ``y-axis`` data points as numbers which appear just above the top of each column in a column chart.
The next section will describe how these methods work.

The call to :py:meth:`.Chart2.print_chart_types` at the end of :py:meth:`.Chart2.insert_chart` could be commented out since it's a diagnostic check.
It prints the names of the chart types used by the template.

.. _ch28_accessing_modifing_chart_el:

28.3 Accessing and Modifying Chart Elements
===========================================

Almost every aspect of a chart can be adjusted, including such things as its color scheme, the fonts, the scaling of the axes, the positioning of the legend, axis labels, and titles.
It's also possible to augment charts with regression line details, error bars, and additional graphs.

These elements are located in a number of different places in the hierarchy of services accessible through the |ChartDocument2|_ service.
A simplified version of this hierarchy is shown in :numref:`ch28fig_chart_doc_hirarchy`.

..
    figure 9

.. cssclass:: diagram invert

    .. _ch28fig_chart_doc_hirarchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/206399293-b5f59e1c-c25c-4f93-970d-a8016dc8d9ef.png
        :alt: The Hierarchy of Services Below ChartDocument
        :figclass: align-center

        :The Hierarchy of Services Below |ChartDocument2|_.

There is more information about the |Diagram2|_, CoordinateSystem_, ChartType_, and DataSeries_ services as this section progresses,
but :numref:`ch28fig_chart_doc_hirarchy` indicates that |Diagram2|_ manages the legend, floor and chart wall,
CoordinateSystem_ is in charge of the axes, and the data points are manipulated via DataSeries_.

The ``1`` and ``*`` in :numref:`ch28fig_chart_doc_hirarchy` indicate that a diagram may utilize multiple coordinate systems,
that a single coordinate system may display multiple chart types, and a single chart type can employ many data series.
Fortunately, this generality isn't often needed for the charts created by :py:meth:`.Chart2.insert_chart`.
In particular, the chart diagram only uses a single coordinate system and a single chart type (most of the time).

.. _ch28_accessing_diagram:

28.3.1 Accessing the Diagram
----------------------------

A chart's Diagram service is easily reached by calling ``ChartDocument.getFirstDiagram()``, which returns a reference to the diagram's |XDiagram2|_ interface:

|XDiagram2|_ contains several useful methods (:abbreviation:`i.e.` ``getLegend()``, ``getWall()``, ``getFloor()``),
and its services hold many properties (:abbreviation:`i.e.` ``StartingAngle`` used in pie charts and ``RotationVertical`` for 3D charts).
This is summarized by :numref:`ch28fig_diagram_srv`.

..
    figure 10

.. cssclass:: diagram invert

    .. _ch28fig_diagram_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/206402610-767ac2a2-4932-4e6b-ad16-c11c7953081c.png
        :alt: The Diagram Service.
        :figclass: align-center

        :The |Diagram2|_ Service.

:py:meth:`.Chart2.set_background_colors` changes the background and wall colors of the chart through the |ChartDocument2|_ and |Diagram2|_ services:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def set_background_colors(
            chart_doc: XChartDocument, bg_color: mColor.Color, wall_color: mColor.Color
        ) -> None:
            try:
                if int(bg_color) > 0:
                    bg_ps = chart_doc.getPageBackground()
                    # Props.show_props("Background", bg_ps)
                    Props.set(
                        bg_ps, FillBackground=True, FillStyle=FillStyle.SOLID, FillColor=int(bg_color)
                    )

                if int(wall_color) > 0:
                    diagram = chart_doc.getFirstDiagram()
                    wall_ps = diagram.getWall()
                    # Props.show_props("Wall", wall_ps)
                    Props.set(
                        wall_ps, FillBackground=True, FillStyle=FillStyle.SOLID, FillColor=int(wall_color)
                    )
            except Exception as e:
                raise ChartError("Error setting background colors") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The chart background is manipulated with a property set accessible through ``XChartDocument.getPageBackground()``, while the wall is reached with ``XDiagram.getWall()``.

The documentation for the ``getPageBackground()`` and ``getWall()`` methods doesn't list the contents of their property sets,
so the easiest way of finding out what's available is by calling :py:meth:`.Props.show_props`. Two ``show_props()`` calls are commented out in the code above.

Most chart services inherit a mix of four property classes:

.. cssclass:: ul-list

    - `com.sun.star.style.CharacterProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html>`_
    - `com.sun.star.style.ParagraphProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html>`_
    - `com.sun.star.drawing.LineProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineProperties.html>`_
    - `com.sun.star.drawing.FillProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1FillProperties.html>`_

Since ``getWall()`` and ``getPageBackground()`` both deal with areas in the chart, their properties come from the ``FillProperties`` class.

.. _ch28_accessing_coord_sys:

28.3.2 Accessing the Coordinate System
--------------------------------------

:numref:`ch28fig_diagram_srv` shows that the diagram's coordinate systems are reached through ``XCoordinateSystemContainer.getCoordinateSystems()``.
:py:meth:`.Chart2.get_coord_system` assumes that the programmer only wants the first coordinate system:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def get_coord_system(chart_doc: XChartDocument) -> XCoordinateSystem:
            try:
                diagram = chart_doc.getFirstDiagram()
                coord_sys_con = Lo.qi(XCoordinateSystemContainer, diagram, True)
                coord_sys = coord_sys_con.getCoordinateSystems()
                if coord_sys:
                    if len(coord_sys) > 1:
                        Lo.print(f"No. of coord systems: {len(coord_sys)}; using first.")
                return coord_sys[0]  # will raise error if coord_sys is empyt or none
            except Exception as e:
                raise ChartError("Error unable to get coord_system") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The CoordinateSystem_ service is employed to access the chart's axes and its chart type (or types), as in :numref:`ch28fig_coordinate_system_service`.

..
    figure 11

.. cssclass:: diagram invert

    .. _ch28fig_coordinate_system_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/206425097-aac4e391-c6be-464b-96d4-40fd12a0e072.png
        :alt: The CoordinateSystem Service
        :figclass: align-center

        :The CoordinateSystem_ Service.

The Axis_ service is described when we look at methods for adjusting axis properties.

.. _ch28_accessing_chart_type:

28.3.3 Accessing the Chart Type
-------------------------------

:numref:`ch28fig_coordinate_system_service` shows that the chart types in a coordinate system are reached through ``XChartTypeContainer.getChartTypes()``.
:py:meth:`.Chart2.get_chart_type` assumes the programmer only wants the first chart type in the array:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def get_chart_type(cls, chart_doc: XChartDocument) -> XChartType:
            try:
                chart_types = cls.get_chart_types(chart_doc)
                return chart_types[0]
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error getting chart type") from e

        @classmethod
        def get_chart_types(cls, chart_doc: XChartDocument) -> Tuple[XChartType, ...]:
            try:
                coord_sys = cls.get_coord_system(chart_doc)
                ct_con = Lo.qi(XChartTypeContainer, coord_sys, True)
                result = ct_con.getChartTypes()
                if result is None:
                    raise UnKnownError("None Value: getChartTypes() returned a value of None")
                return result
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error getting chart types") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:numref:`ch28fig_chart_type_srv` shows the main components of the ChartType_ service.

..
    figure 12

.. cssclass:: diagram invert

    .. _ch28fig_chart_type_srv:
    .. figure:: https://user-images.githubusercontent.com/4193389/206427238-b3258dcc-1982-4ebe-92ac-b5f64f73aadf.png
        :alt: The ChartType Service
        :figclass: align-center

        :The ChartType_ Service.

Somewhat surprisingly, the ChartType_ service isn't the home for chart type related properties;
instead XChartType_ contains methods for examining chart type "roles", which is described later.
One useful features of XChartType_ is ``getChartType()`` which returns the type as a string.

The CandleStickChartType_ service inherits ChartType_, and contains properties related to stock charts.

.. _ch28_accessing_data_series:

28.3.4 Accessing the Data Series
--------------------------------

:numref:`ch28fig_chart_type_srv` shows that the data series for a chart type is accessed via ``XDataSeriesContainer.getDataSeries()``.
This is implemented by :py:meth:`.Chart2.get_data_series`:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def get_data_series(
            cls, chart_doc: XChartDocument, chart_type: ChartTypeNameBase | str = ""
        ) -> Tuple[XDataSeries, ...]:
            try:
                if chart_type:
                    xchart_type = cls.find_chart_type(chart_doc, chart_type)
                else:
                    xchart_type = cls.get_chart_type(chart_doc)
                ds_con = Lo.qi(XDataSeriesContainer, xchart_type, True)
                return ds_con.getDataSeries()
            except Exception as e:
                raise ChartError("Error getting chart data series") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The DataSeries_ service is one of the more complex parts of the Chart2 module because of its support for several important interfaces.
They will not all be explained just yet; :numref:`ch28fig_data_series_xdata_series` focuses on the XDataSeries_ interface.

..
    figure 13

.. cssclass:: diagram invert

    .. _ch28fig_data_series_xdata_series:
    .. figure:: https://user-images.githubusercontent.com/4193389/206428966-023cbd67-f3fc-4fcf-a23b-c6b0be204ac7.png
        :alt: The DataSeries Service and XDataSeries Interface
        :figclass: align-center

        :The DataSeries_ Service and XDataSeries_ Interface.

A DataSeries_ represents a series of data points in the chart.
Changes to the look of these data points (:abbreviation:`i.e.` adding numbers next to the points, or changing their shape and color) can be done in two ways.
A data series as a whole maintains a set of properties, most of which are inherited from the DataPointProperties_ class.
Typical DataPointProperties_ values are ``Color``, ``Shape``, ``LineWidth``.

It's also possible to adjust point properties on an individual basis by accessing a particular data point by calling ``XDataSeries.getDataPointByIndex()``.
As the method name suggests, this requires an index value for the point, which can be a little tricky to determine.

Now we can explain the second of the two chart changing methods called at the end of :py:meth:`.Chart2.insert_chart`: :py:meth:`.Chart2.set_data_point_labels`,
which switches on the displaying of the ``y-axis`` data points as numbers.
The call is:

.. tabs::

    .. code-tab:: python

        # part of Chart2.insert_chart()...
        cls.set_data_point_labels(chart_doc, DataPointLabelTypeKind.NUMBER)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.set_data_point_labels` uses :py:meth:`.Chart2.get_data_series` described above, which returns an array of all the data series used in the chart.
:py:meth:`~.Chart2.set_data_point_labels` iterates through the array and manipulates the ``Label`` property for each series.
In other words, it modifies each data series 

.. tabs::

    .. code-tab:: python

        # 
        @classmethod
        def set_data_point_labels(
            cls, chart_doc: XChartDocument, label_type: DataPointLabelTypeKind
        ) -> None:
            try:
                data_series_arr = cls.get_data_series(chart_doc=chart_doc)
                for data_series in data_series_arr:
                    dp_label = cast(DataPointLabel, Props.get_property(data_series, "Label"))
                    dp_label.ShowNumber = False
                    dp_label.ShowCategoryName = False
                    dp_label.ShowLegendSymbol = False
                    if label_type == DataPointLabelTypeKind.NUMBER:
                        dp_label.ShowNumber = True
                    elif label_type == DataPointLabelTypeKind.PERCENT:
                        dp_label.ShowNumber = True
                        dp_label.ShowNumberInPercent = True
                    elif label_type == DataPointLabelTypeKind.CATEGORY:
                        dp_label.ShowCategoryName = True
                    elif label_type == DataPointLabelTypeKind.SYMBOL:
                        dp_label.ShowLegendSymbol = True
                    elif label_type == DataPointLabelTypeKind.NONE:
                        pass
                    else:
                        raise UnKnownError("label_type is of unknow type")

                    Props.set_property(data_series, "Label", dp_label)
            except ChartError:
                raise
            except Exception as e:
                raise ChartError("Error setting data point labels") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    :py:class:`~.kind.data_point_label_type_kind.DataPointLabelTypeKind`

The ``Label`` DataSeries_ property is inherited from DataPointProperties_.
``Label`` is of type DataPointLabel_ which maintains four 'show' boolean values for displaying the number and other kinds of information next to the data point.
Depending on the ``label_type`` value passed to :py:meth:`.Chart2.set_data_point_labels`, one or more of these boolean values are set and the ``Label`` property updated.



.. |ChartDocument2| replace:: ChartDocument
.. _ChartDocument2: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1ChartDocument.html

.. |XChartDocument2| replace:: XChartDocument
.. _XChartDocument2: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XChartDocument.html

.. |Diagram2| replace:: Diagram
.. _Diagram2: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1Diagram.html

.. |XDiagram2| replace:: XDiagram
.. _XDiagram2: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XDiagram.html

.. |ug| replace:: Calc User Guide
.. _ug: https://books.libreoffice.org/en/CG74/CG74.html

.. |ug_ch03| replace:: chapter 3
.. _ug_ch03: https://books.libreoffice.org/en/CG74/CG7403-ChartsAndGraphs.html

.. _Axis: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1Axis.html
.. _CandleStickChartType: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1CandleStickChartType.html
.. _ChartType: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1ChartType.html
.. _CoordinateSystem: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1CoordinateSystem.html
.. _DataPointLabel: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1chart2_1_1DataPointLabel.html
.. _DataPointProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1DataPointProperties.html
.. _DataSeries: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1DataSeries.html
.. _TableChart: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableChart.html
.. _TableCharts: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1table_1_1TableCharts.html
.. _TabularDataProviderArguments: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1chart2_1_1data_1_1TabularDataProviderArguments.html
.. _XChartType: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XChartType.html
.. _XChartTypeManager: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XChartTypeManager.html
.. _XDataProvider: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1data_1_1XDataProvider.html
.. _XDataSeries: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XDataSeries.html
.. _XDataSource: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1data_1_1XDataSource.html
.. _XEmbeddedObjectSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XEmbeddedObjectSupplier.html
.. _XTableChart: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1table_1_1XTableChart.html

.. spelling:word-list::
    Donut
