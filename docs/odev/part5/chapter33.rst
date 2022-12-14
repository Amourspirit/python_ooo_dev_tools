.. _ch33:

*******************************************
Chapter 33. Using Charts in Other Documents
*******************************************

.. topic:: Overview

    Copy-and-Paste Dispatches; Adding a Chart to a Text Document; Adding a Chart to a Slide Document; Saving the Chart as an Image

    Examples: |slide_chart|_ and |text_chart|_

This chapter describes two examples, |slide_chart|_ and |text_chart|_, which illustrate a copy-and-paste approach to adding charts to text and slide documents.

|slide_chart|_ also shows how to save a chart as an image file.

There are three tricky aspects to coding with the clipboard.

1. | Clipboard manipulation requires the use of UNO dispatch commands (namely ``.uno:Copy`` and ``.uno:Paste``).
   | The function that sends out a dispatch command returns immediately, but Office may take a few milliseconds to process the request.
   | This often introduces a timing issue because the next step in the program should be delayed in order to give the dispatch time to be processed.
2. | Another timing problem is that dispatch commands assume that the Office application window is visible and active.
   | Office can be made the focus of the OS, but the few milliseconds between achieving this and sending the dispatch may be enough for another OS process to pull the focus away from Office.
   | The dispatch will then arrive at the wrong process, with unpredictable results.
3. | Clipboard programming usually involves two documents.
   | In this case, the spreadsheet that generates the chart, and the text or slide document that receives the image.
   | One of the limitations of |odev| classes is that many methods assume only one document is open at a time.
   | To be safe, the program should close the spreadsheet before the text or slide document is opened or created.

Bearing all these issues in mind, is there a better approach for placing a chart in a non-spreadsheet document?
The answer should be OLE (Object Linking and Embedding), which is Windows' solution to this very problem.
Unfortunately, Hove not figured out how to get this mechanism to work via the API.
Looking back through old forum posts, it appears that OLE didn't work in the old chart module either.

.. todo:: 

    Chapter 32, insert link to chapter 43

Clipboard programming is looked at again in Chapter 43, Python's copy-and-paste API is considered, and describe more general examples using Writer, Calc, Impress, and Base documents.

.. _ch33_adding_chart_txt_doc:

33.1 Adding a Chart to a Text Document
======================================

|text_chart_py|_ generates a column chart using the "Sneakers Sold this Month" table from |ods_doc|, copies it to the clipboard, and closes the spreadsheet.
Then a text document is created, and the chart image is pasted into it, resulting in the page shown in :numref:`ch33fig_text_doc_chart`.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch33fig_text_doc_chart:
    .. figure:: https://user-images.githubusercontent.com/4193389/207697020-04cc5764-960d-4145-857d-4b43cc6ff013.png
        :alt: Text Document with a Chart Image
        :figclass: align-center

        :Text Document with a Chart Image.

|text_chart_py|_ relies on the Write support class from :ref:`part02`. The "Hello LibreOffice." paragraph is written into a new document,
then the chart is pasted in from the clipboard, a paragraph ("Figure 1. Sneakers Column Chart.") added as a figure legend,
and another paragraph ("Some more text…") added at the end for good measure.

``main()`` is:

.. tabs::

    .. code-tab:: python

        # TextChart.main() of text_chart.py
        def main(self) -> None:
            loader = Lo.load_office(Lo.ConnectPipe())

            try:
                has_chart = self._make_col_chart(loader)

                doc = Write.create_doc(loader)
                # to make the construction visible
                GUI.set_visible(is_visible=True, odoc=doc)

                cursor = Write.get_cursor(doc)
                # make sure at end of doc before appending
                cursor.gotoEnd(False)

                Write.append_para(cursor=cursor, text="Hello LibreOffice.\n")

                if has_chart:
                    Lo.delay(1_000)
                    Lo.dispatch_cmd(GlobalEditDispatch.PASTE)

                Write.append_para(cursor=cursor, text="Figure 1. Sneakers Column Chart.\n")
                Write.style_prev_paragraph(
                    cursor=cursor, prop_val=ParagraphAdjust.CENTER, prop_name="ParaAdjust"
                    )

                Write.append_para(cursor=cursor, text="Some more text...\n")

                Lo.delay(2000)
                msg_result = MsgBox.msgbox(
                    "Do you wish to close document?",
                    "All done",
                    boxtype=MessageBoxType.QUERYBOX,
                    buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
                )
                if msg_result == MessageBoxResultsEnum.YES:
                    Lo.close_doc(doc=doc, deliver_ownership=True)
                    Lo.close_office()
                else:
                    print("Keeping document open")
            except Exception:
                Lo.close_office()
                raise

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

It's important that the text document is visible and in focus, so :py:meth:`.GUI.set_visible` is called after the document's creation.
There's also a call to :py:meth:`.Lo.delay` before the paste (``Lo.dispatch_cmd(GlobalEditDispatch.PASTE)``) to ensure that earlier text writes have time to finish.

.. _ch33_making_chart:

33.1.1 Making the Chart
-----------------------

|text_chart_py|_ uses ``_make_col_chart()`` to generate a chart from a spreadsheet:

.. tabs::

    .. code-tab:: python

        # TextChart._make_col_chart() of text_chart.py
        def _make_col_chart(self, loader: XComponentLoader) -> bool:
            ssdoc = Calc.open_doc(fnm=self._data_fnm, loader=loader)
            try:
                GUI.set_visible(is_visible=True, odoc=ssdoc)  # or selection is not copied
                sheet = Calc.get_sheet(doc=ssdoc, index=0)

                range_addr = Calc.get_address(sheet=sheet, range_name="A2:B8")
                chart_doc = Chart2.insert_chart(
                    sheet=sheet,
                    cells_range=range_addr,
                    cell_name="C3",
                    width=15,
                    height=11,
                    diagram_name=ChartTypes.Column.TEMPLATE_STACKED.COLUMN,
                )

                Chart2.set_title(
                    chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A1")
                )
                Chart2.set_x_axis_title(
                    chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A2")
                )
                Chart2.set_y_axis_title(
                    chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B2")
                )
                Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
                Lo.delay(1_000)
                Chart2.copy_chart(ssdoc=ssdoc, sheet=sheet)
                return True
            except Exception as e:
                Lo.print("Error making col chart")
                Lo.print(f"  {e}")
            finally:
                Lo.close_doc(doc=ssdoc)
            return False

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``_make_col_chart()`` uses the table from |ods_doc| shown in :numref:`ch33fig_sneakers_month_tbl` to generate the chart in ::numref:`ch33fig_text_doc_chart` .

..
    figure 2

.. cssclass:: screen_shot invert

    .. _ch33fig_sneakers_month_tbl:
    .. figure:: https://user-images.githubusercontent.com/4193389/207700554-1331df5f-1881-45dc-af53-834109345903.png
        :alt: The Sneakers Sold this Month Table
        :figclass: align-center

        :The "Sneakers Sold this Month" Table.

The only new feature in ``_make_col_chart()`` is the call to :py:meth:`.Chart2.copy_chart` which copies the chart to the clipboard.

Two easy to overlook parts of ``_make_col_chart()`` are the call to :py:meth:`.GUI.set_visible`, which makes the spreadsheet and chart visible and active, and the call to
:py:meth:`.Lo.delay` before :py:meth:`.Chart2.copy_chart`.
This ensures that there's enough time for the graph to be drawn before the ``.uno:Copy`` dispatch.

Also note that the spreadsheet is closed before ``_make_col_chart()`` returns.
This stops the subsequent creation of the text document back in ``main()`` from being possibly affected by an open spreadsheet.

.. _ch33_copying_chart:

33.1.2 Copying a Chart
----------------------

:py:meth:`.Chart2.copy_chart` obtains a reference to the chart as an XShape_, which makes it possible to select it with an XSelectionSupplier_.
This selection is used automatically as the data for the ``.uno:Copy`` dispatch.
:py:meth:`~.Chart2.copy_chart` is:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def copy_chart(cls, ssdoc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
            try:
                chart_shape = cls.get_chart_shape(sheet=sheet)
                doc = Lo.qi(XComponent, ssdoc, True)
                supp = GUI.get_selection_supplier(doc)
                supp.select(chart_shape)
                Lo.dispatch_cmd("Copy")
            except Exception as e:
                raise ChartError("Error in attempt to copy chart") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Every spreadsheet is also a draw page, so the Spreadsheet_ service has an XDrawPageSupplier_ interface, and its ``getDrawPage()`` method returns an XDrawPage_ reference.
For example:

.. tabs::

    .. code-tab:: python

        # part of Chart2.get_chart_shape(); see below
        page_supp = Lo.qi(XDrawPageSupplier, sheet, True)
        draw_page = page_supp.getDrawPage()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The shapes in a draw page can be accessed by index.
Also each shape has a ``CLSID`` property which can be used to identify 'special' shapes representing math formulae or charts.
The search for a chart shape is coded as:

.. tabs::

    .. code-tab:: python

        # part of Chart2.get_chart_shape(); see below
        num_shapes = draw_page.getCount()
        chart_classid = Lo.CLSID.CHART.value
        for i in range(num_shapes):
            try:
                shape = mLo.Lo.qi(XShape, draw_page.getByIndex(i), True)
                classid = str(Props.get(shape, "CLSID")).lower()
                if classid == chart_classid:
                    break
            except Exception:
                shape = None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

These two pieces of code are combined in :py:meth:`.Chart2.get_chart_shape`:

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @staticmethod
        def get_chart_shape(sheet: XSpreadsheet) -> XShape:
            shape = None
            try:
                page_supp = Lo.qi(XDrawPageSupplier, sheet, True)
                draw_page = page_supp.getDrawPage()
                num_shapes = draw_page.getCount()
                chart_classid = Lo.CLSID.CHART.value
                for i in range(num_shapes):
                    try:
                        shape = Lo.qi(XShape, draw_page.getByIndex(i), True)
                        classid = str(Props.get(shape, "CLSID")).lower()
                        if classid == chart_classid:
                            break
                    except Exception:
                        shape = None
                        # continue on, just because got an error does not mean shape will not be found
            except Exception as e:
                raise ShapeError("Error getting shape from sheet") from e
            if shape is None:
                raise ShapeMissingError("Unalbe to find Chart Shape")
            return shape

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch33_adding_chart_slide_doc:

33.2 Adding a Chart to a Slide Document
=======================================

|slide_chart|_ generates the same column chart as |text_chart|_, using almost the same version of ``_make_col_chart()``.
After the chart has been copied to the clipboard and the spreadsheet closed, a slide document is created and the chart pasted onto the first slide.
The chart appears in the center of the slide by default, but is moved down to make room for some text.
The end result is shown in :numref:`ch33fig_slide_chart`.

..
    figure 3

.. cssclass:: screen_shot invert

    .. _ch33fig_slide_chart:
    .. figure:: https://user-images.githubusercontent.com/4193389/207704581-51673b9c-a33e-4fcc-b32c-f319472bc0ce.png
        :alt: Slide Document with a Chart Shape.
        :figclass: align-center
        :width: 550px

        :Slide Document with a Chart Shape.

The ``main()`` function of |slide_chart_py|_ is:

.. tabs::

    .. code-tab:: python

        # SlideChart.main() of slide_chart.py
        def main(self) -> None:
            loader = Lo.load_office(Lo.ConnectPipe())

            try:
                has_chart = self._make_col_chart(loader)

                doc = Draw.create_impress_doc(loader)
                # to make the construction visible
                GUI.set_visible(is_visible=True, odoc=doc)

                # access first page.
                slide = Draw.get_slide(doc=doc, idx=0)
                body = Draw.bullets_slide(slide=slide, title="Sneakers Are Selling!")
                Draw.add_bullet(
                    bulls_txt=body, level=0, text="Sneaker profits have increased"
                )

                if has_chart:
                    Lo.delay(1_000)
                    Lo.dispatch_cmd(GlobalEditDispatch.PASTE)

                try:
                    ole_shape = Draw.find_shape_by_type(
                        slide=slide, shape_type=DrawingNameSpaceKind.OLE2_SHAPE
                    )
                    slide_size = Draw.get_slide_size(slide)
                    shape_size = Draw.get_size(ole_shape)
                    shape_pos = Draw.get_position(ole_shape)

                    y = slide_size.Height - shape_size.Height - 20
                    # move pic down
                    Draw.set_position(shape=ole_shape, x=shape_pos.X, y=y)
                except mEx.ShapeMissingError:
                    Lo.print("Did not find shape, unable to set size and position")

                Lo.delay(2000)
                msg_result = MsgBox.msgbox(
                    "Do you wish to close document?",
                    "All done",
                    boxtype=MessageBoxType.QUERYBOX,
                    buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
                )
                if msg_result == MessageBoxResultsEnum.YES:
                    Lo.close_doc(doc=doc, deliver_ownership=True)
                    Lo.close_office()
                else:
                    print("Keeping document open")
            except Exception:
                Lo.close_office()
                raise

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The chart is pasted into the slide as an ``OLE2Shape`` object, which allows it to be found by :py:meth:`.Draw.find_shape_by_type`.
The shape is moved down the slide by calculating a new (``x``, ``y``) coordinate for its top-left corner, and calling :py:meth:`.Draw.set_position`.

.. _ch33_save_chart_img:

33.3 Saving the Chart as an Image
=================================

The only change to ``_make_col_chart()`` in |slide_chart_py|_ is the addition of:

.. tabs::

    .. code-tab:: python

        # in _make_col_chart() of slide_chart.py
        try:
            ImagesLo.save_graphic(
                pic=Chart2.get_chart_image(sheet),
                fnm=Path(self._out_dir, "chartImage.png")
            )
        except mEx.ImageError:
            pass

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This saves the chart as a PNG image, which can be loaded by other applications.

:py:meth:`.ImagesLo.save_graphic` accepts a XGraphic_ argument and filename:

.. tabs::

    .. code-tab:: python

        # in ImageLo class
        @staticmethod
        def save_graphic(pic: XGraphic, fnm: PathOrStr, im_format: str = "") -> None:

            Lo.print(f"Saving graphic in '{fnm}'")

            try:
                if pic is None:
                    raise TypeError("Expected pic to be XGraphic instance but got None")
                if not im_format:
                    im_format = Info.get_ext(fnm)
                    if not im_format:
                        raise ValueError(
                            "Unable to get image format from fnm. Does fnm have an file extension such as myfile.png?"
                        )
                    im_format = im_format.lower()

                gprovider = Lo.create_instance_mcf(
                    XGraphicProvider, "com.sun.star.graphic.GraphicProvider", raise_err=True
                )

                png_props = Props.make_props(
                    URL=mFileIO.FileIO.fnm_to_url(fnm), MimeType=f"image/{im_format}"
                )

                gprovider.storeGraphic(pic, png_props)
            except Exception as e:
                raise ImageError(f'Error saving graphic for "{fnm}') from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.get_chart_image` finds the chart in the spreadsheet and returns it as a XGraphic_ object.

.. tabs::

    .. code-tab:: python

        # in Chart2 class
        @classmethod
        def get_chart_image(cls, sheet: XSpreadsheet) -> XGraphic:
            try:
                chart_shape = cls.get_chart_shape(sheet)

                graphic = Lo.qi(
                    XGraphic,
                    Props.get(chart_shape, "Graphic"),
                    True
                )

                tmp_fnm = FileIO.create_temp_file("png")
                ImagesLo.save_graphic(pic=graphic, fnm=tmp_fnm, im_format="png")
                im = ImagesLo.load_graphic_file(tmp_fnm)
                FileIO.delete_file(tmp_fnm)
                return im
            except Exception as e:
                raise ChartError("Error getting chart image") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Chart2.get_chart_image` finds the chart in the sheet by using :py:meth:`.Chart2.get_chart_shape` described earlier.
The shape is cast to an Office graphics object, of type XGraphic_.

:py:meth:`~.Chart2.get_chart_image` then creates a temporary file to store the XGraphic_ image, which is immediately re-loaded.

.. |ods_doc| replace:: ``chartsData.ods``

.. |slide_chart| replace:: Slide Chart
.. _slide_chart: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/chart2/slide_chart

.. |slide_chart_py| replace:: slide_chart.py
.. _slide_chart_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/chart2/slide_chart/slide_chart.py

.. |text_chart| replace:: Text Chart
.. _text_chart: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/chart2/text_chart

.. |text_chart_py| replace:: text_chart.py
.. _text_chart_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/chart2/text_chart/text_chart.py

.. _Spreadsheet: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1Spreadsheet.html
.. _XDrawPage: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPage.html
.. _XDrawPageSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPageSupplier.html
.. _XGraphic: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1graphic_1_1XGraphic.html
.. _XSelectionSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1view_1_1XSelectionSupplier.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
