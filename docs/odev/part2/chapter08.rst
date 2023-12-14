.. _ch08:

**************************
Chapter 8. Graphic Content
**************************

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

.. topic:: Overview

    Graphics; Linked Images/Shapes

    Examples: |build_doc|_ and |extract_graphics|_.

:ref:`ch07` looked at several forms of text document content (e.g. text frames, math formulae, text fields and tables, and bookmarks),
as indicated by :numref:`ch08fig_text_content_serv_subs`.
However the different ways of adding graphical content (corresponding to the services highlighted)
are the focus of this chapter

.. cssclass:: diagram invert

    .. _ch08fig_text_content_serv_subs:
    .. figure:: https://user-images.githubusercontent.com/4193389/185280015-9f384d68-0014-4e21-9f66-224fbea61f5f.png
        :alt: Diagram of The TextContent Service and Some Sub-classes.
        :figclass: align-center

        :The TextContent Service and Some Sub-classes.

.. _ch08_link_graphic:

8.1 Linking a Graphic Object to a Document
==========================================

Adding an image to a text document follows the same steps as other text content, as shown in :py:meth:`.Write.add_image_link`:

.. tabs::

    .. code-tab:: python

        # in Write class
        @classmethod
        def add_image_link(
            cls, doc: XTextDocument, cursor: XTextCursor, fnm: PathOrStr, width: int = 0, height: int = 0
        ) -> bool:
            cargs = CancelEventArgs(Write.add_image_link.__qualname__)
            cargs.event_data = {
                "doc": doc,
                "cursor": cursor,
                "fnm": fnm,
                "width": width,
                "height": height,
            }
            _Events().trigger(WriteNamedEvent.IMAGE_LINK_ADDING, cargs)
            if cargs.cancel:
                return False

            fnm = cargs.event_data["fnm"]
            width = cargs.event_data["width"]
            height = cargs.event_data["height"]

            try:
                tgo = Lo.create_instance_msf(XTextContent, "com.sun.star.text.TextGraphicObject")
                if tgo is None:
                    raise mEx.CreateInstanceMsfError(XTextContent, "com.sun.star.text.TextGraphicObject")

                props = Lo.qi(XPropertySet, tgo, True)
                props.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER)
                props.setPropertyValue("GraphicURL", FileIO.fnm_to_url(fnm))

                # optionally set the width and height
                if width > 0 and height > 0:
                    props.setPropertyValue("Width", width)
                    props.setPropertyValue("Height", height)

                # append image to document, followed by a newline
                cls._append_text_content(cursor, tgo)
                cls.end_line(cursor)
            except CreateInstanceMsfError:
                raise
            except MissingInterfaceError:
                raise
            except Exception as e:
                raise Exception(f"Insertion of graphic in '{fnm}' failed:") from e
            _Events().trigger(WriteNamedEvent.IMAGE_LINK_ADDED, EventArgs.from_args(cargs))
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The TextGraphicObject_ service doesn't offer a ``XTextGraphicObject`` interface, so :py:meth:`.Lo.create_instance_msf`
returns an XTextContent_.

The interface is also converted to XPropertySet_ because several properties need to be set.
The frame is anchored, and the image's filename is assigned to ``GraphicURL`` (after being changed into a URL).

The image's size on the page depends on the dimensions of its enclosing frame, which are set in the "Width" and "Height" properties:

.. tabs::

    .. code-tab:: python

        props.setPropertyValue("Width", 4_500) # 45 mm width
        props.setPropertyValue("Height", 4_000) # 40 mm height

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The values are in ``1/100`` mm units, so ``4500`` is ``45`` mm or ``4.5 cm``.

If these properties aren't explicitly set then the frame size defaults to being the width and height of the image.

In more realistic code, the width and height properties would be calculated as some scale factor of the image's size,
as measured in ``1/100 mm`` units not pixels.
These dimensions are available if the image file is loaded as an XGraphic_ object, as shown in :py:meth:`.ImagesLo.get_size_100mm`:

.. tabs::

    .. code-tab:: python

        # in ImagesLo class
        @classmethod
        def get_size_100mm(cls, im_fnm: PathOrStr) -> Size:
            graphic = cls.load_graphic_file(im_fnm)
            return Props.get_property(prop_set=graphic, name="Size100thMM")

        @staticmethod
        def load_graphic_link(graphic_link: object) -> XGraphic:
            gprovider = Lo.create_instance_mcf(
                XGraphicProvider, "com.sun.star.graphic.GraphicProvider", raise_err=True
            )

            xprops = Lo.qi(XPropertySet, graphic_link, True)

            try:
                gprops = Props.make_props(URL=str(xprops.getPropertyValue("GraphicURL")))
                return gprovider.queryGraphic(gprops)
            except Exception as e:
                raise Exception(f"Unable to retrieve graphic") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Displaying the image at a scaled size is possible by combining :py:meth:`.ImagesLo.get_size_100mm` and :py:meth:`.Write.add_image_link`:

.. tabs::

    .. code-tab:: python

        img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)

        # enlarge by 1.5x
        h = round(img_size.Height * 1.5)
        w = round(img_size.Width * 1.5)
        Write.add_image_link(doc, cursor, im_fnm, w, h)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A possible drawback of :py:meth:`.Write.add_image_link` is that the document only contains a link to the image.
This becomes an issue if you save the document in a format other than ``.odt``.
In particular, when saved as a Word ``.doc`` file, the link is lost.

.. _ch08_add_graphic_shape:

8.2 Adding a Graphic to a Document as a Shape
=============================================

An alternative to inserting a graphic as a link is to treat it as a shape.
Shapes will be discussed at length in :ref:`part03`, so I won't go into much detail about them here.
One difference between a graphic link and shape is that shapes can be rotated.

Shapes can be created using the com.sun.star.text.Shape_ service, com.sun.star.drawing.Shape_,
or one of its sub-classes, while ``XDrawPageSupplier.getDrawPage()`` accesses the shapes in a document.

The shape hierarchy is quite extensive (i.e. there are many kinds of shape), so only the parts used here are shown in :numref:`ch08fig_shape_hierachy_parts`:

.. cssclass:: diagram invert

    .. _ch08fig_shape_hierachy_parts:
    .. figure:: https://user-images.githubusercontent.com/4193389/185285831-8f1fe982-06ea-4e17-a64d-c76ea74662fe.png
        :alt: Diagram of Part of the Shape Hierarchy
        :figclass: align-center

        :Part of the Shape Hierarchy.

In :numref:`ch08fig_shape_hierachy_parts`, "(text) Shape" refers to the com.sun.star.text.Shape_ service, while "(drawing) Shape"
is com.sun.star.drawing.Shape_.

The examples use GraphicObjectShape_ (:py:meth:`.Write.add_image_shape`) to create a shape containing an image, and LineShape_ (:py:meth:`.Write.add_line_divider`) to add a line to the document.

The XShapeDescriptor_ interface in com.sun.star.drawing.Shape_ is a useful way to obtain the name of a shape service.

.. _ch08_create_img_shape:

8.2.1 Creating an Image Shape
-----------------------------

The |build_doc|_ example adds an image shape to the document by calling :py:meth:`.Write.add_image_shape`:

.. tabs::

    .. code-tab:: python

        # code fragment from build doc
        # add image as shape to page
        cursor.append_line("Image as a shape: ")
        cursor.add_image_shape(fnm=im_fnm)
        cursor.end_paragraph()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.WriteTextViewCursor.add_image_shape` invokes :py:meth:`.Write.add_image_shape`, which comes in two versions: with and without width and height arguments.
A shape with no explicitly set width and height properties is rendered as a miniscule image (about 1 mm wide).
Call me old-fashioned, but I want to see the graphic, so :py:meth:`.Write.add_image_shape` calculates the picture's
dimensions if none are supplied by the user.

Another difference between image shape and image link is how the content's ``GraphicURL`` property is employed.
The image link version contains its URL, while the image shape's ``GraphicURL`` stores its bitmap as a string.

The code for :py:meth:`.Write.add_image_shape`:

.. tabs::

    .. code-tab:: python

        # in Write class
        @classmethod
        def add_image_shape(
            cls, cursor: XTextCursor, fnm: PathOrStr, width: int = 0, height: int = 0
        ) -> bool:
            cargs = CancelEventArgs(Write.add_image_shape.__qualname__)
            cargs.event_data = {
                "cursor": cursor,
                "fnm": fnm,
                "width": width,
                "height": height,
            }
            _Events().trigger(WriteNamedEvent.IMAGE_SHAPE_ADDING, cargs)
            if cargs.cancel:
                return False

            # get value after event has been raised in case any have been changed.
            fnm = cargs.event_data["fnm"]
            width = cargs.event_data["width"]
            height = cargs.event_data["height"]

            pth = FileIO.get_absolute_path(fnm)

            try:
                if width > 0 and height > 0:
                    im_size = Size(width, height)
                else:
                    im_size = ImagesLo.get_size_100mm(pth)  # in 1/100 mm units
                    if im_size is None:
                        raise ValueError(f"Unable to get image from {pth}")

                # create TextContent for an empty graphic
                gos = Lo.create_instance_msf(
                    XTextContent, "com.sun.star.drawing.GraphicObjectShape", raise_err=True
                )

                bitmap = ImagesLo.get_bitmap(pth)
                if bitmap is None:
                    raise ValueError(f"Unable to get bitmap of {pth}")
                # store the image's bitmap as the graphic shape's URL's value
                Props.set_property(prop_set=gos, name="GraphicURL", value=bitmap)

                # set the shape's size
                xdraw_shape = Lo.qi(XShape, gos, True)
                xdraw_shape.setSize(im_size)

                # insert image shape into the document, followed by newline
                cls._append_text_content(cursor, gos)
                cls.end_line(cursor)
            except ValueError:
                raise
            except CreateInstanceMsfError:
                raise
            except MissingInterfaceError:
                raise
            except Exception as e:
                raise Exception(f"Insertion of graphic in '{fnm}' failed:") from e
            _Events().trigger(WriteNamedEvent.IMAGE_SHAPE_ADDED, EventArgs.from_args(cargs))
            return True

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The image's size is calculated using :py:meth:`.ImagesLo.get_size_100mm` if the user doesn't supply a width and height,
and is used towards the end of the method.

An image shape is created using the GraphicObjectShape_ service, and its XTextContent_ interface is converted to XPropertySet_ for
assigning its properties, and to XShape_ for setting its size (see :numref:`ch08fig_shape_hierachy_parts`).
XShape_ includes a ``setSize()`` method.

.. _ch08_add_oth_graphic:

8.2.2 Adding Other Graphics to the Document
===========================================

The graphic text content can be any sub-class of Shape. In the last section I created a GraphicObjectShape_ service, and accessed its XTextContent_ interface:

.. tabs::

    .. code-tab:: python

        gos = Lo.create_instance_msf(XTextContent, "com.sun.star.drawing.GraphicObjectShape")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

In this section I'll use LineShape:

.. tabs::

    .. code-tab:: python

        ls = Lo.create_instance_msf(XTextContent, "com.sun.star.drawing.LineShape")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The aim is to draw an horizontal line in the document, to act as a divider between paragraphs. The line will be half-a-page wide and centered, like the one in :numref:`ch07fig_graphical_line_divider_ss`.

.. cssclass:: screen_shot invert

    .. _ch07fig_graphical_line_divider_ss:
    .. figure:: https://user-images.githubusercontent.com/4193389/185510906-7bf681f4-4893-4ff9-a721-4dc4d70414e0.png
        :alt: Screen shot of Graphical Line Divider
        :figclass: align-center

        :A Graphical Line Divider.

The difficult part is calculating the width of the line, which should only extend across half the writing width.
This isn't the same as the page width because it doesn't include the left and right margins.

The page and margin dimensions are accessible through the "Standard" page style, as implemented in :py:meth:`.Write.get_page_text_width`:

.. tabs::

    .. code-tab:: python

        # in Write class

        @staticmethod
        def get_page_text_width(text_doc: XTextDocument) -> int:
            props = Info.get_style_props(doc=text_doc, family_style_name="PageStyles", prop_set_nm="Standard")
            if props is None:
                Lo.print("Could not access the standard page style")
                return 0

            try:
                width = int(props.getPropertyValue("Width"))
                left_margin = int(props.getPropertyValue("LeftMargin"))
                right_margin = int(props.getPropertyValue("RightMargin"))
                return width - (left_margin + right_margin)
            except Exception as e:
                Lo.print("Could not access standard page style dimensions")
                Lo.print(f"    {e}")
                return 0

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`~.Write.get_page_text_width` returns the writing width in ``1/100 mm`` units, which is scaled, then passed to :py:meth:`.Write.add_line_divider`:

.. seealso::

    :ref:`ns_units`

.. tabs::

    .. code-tab:: python

        # code fragment in build doc
        text_width = doc.get_page_text_width()
        # scale width by 0.5
        cursor.add_line_divider(line_width=round(text_width * 0.5))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`~.Write.add_line_divider` creates a LineShape_ service with an XTextContent_ interface (see :numref:`ch08fig_shape_hierachy_parts`).
This is converted to XShape_ so its ``setSize()`` method can be passed the line width:

.. tabs::

    .. code-tab:: python

        # in Write class
        @classmethod
        def add_line_divider(cls, cursor: XTextCursor, line_width: int) -> None:
            try:
                ls = Lo.create_instance_msf(XTextContent, "com.sun.star.drawing.LineShape")
                if ls is None:
                    raise CreateInstanceMsfError(XTextContent, "com.sun.star.drawing.LineShape")

                line_shape = Lo.qi(XShape, ls, True)
                line_shape.setSize(Size(line_width, 0))

                cls.end_paragraph(cursor)
                cls._append_text_content(cursor, ls)
                cls.end_paragraph(cursor)

                # center the previous paragraph
                cls.style_prev_paragraph(
                    cursor=cursor, prop_val=ParagraphAdjust.CENTER, prop_name="ParaAdjust"
                )

                cls.end_paragraph(cursor)
            except CreateInstanceMsfError:
                raise
            except MissingInterfaceError:
                raise
            except Exception as e:
                raise Exception("Insertion of graphic line failed") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The centering of the line is achieved by placing the shape in its own paragraph, then using :py:meth:`.Write.style_prev_paragraph` to center it.

.. _ch08_access_link_img:

8.3 Accessing Linked Images and Shapes
======================================

The outcome of running|build_doc|_ is a ``build.odt`` file containing four graphics – two are linked images, one is an image shape, and the other a line shape.

The |extract_graphics|_ example extracts linked graphics from a document, saving them as PNG files.

.. cssclass:: rst-collapse

    .. collapse:: Output:
        :open:

        ::

            No. of text graphics: 2
            Saving graphic in 'C:\Users\user\AppData\Local\Temp\tmpixludwxs\graphics0.png'
            Image size in pixels: 319 X 274
            Saving graphic in 'C:\Users\user\AppData\Local\Temp\tmpixludwxs\graphics1.png'
            Image size in pixels: 319 X 274

            Could not obtain text shapes supplier

            No. of draw shapes: 5
            Shape Name: Shape1
              Type: com.sun.star.drawing.GraphicObjectShape
              Point (mm): [0, 0]
              Size (mm): [61, 58]
            Shape Name: Shape2
              Type: com.sun.star.drawing.LineShape
              Point (mm): [0, 0]
              Size (mm): [88, 0]
            Shapes does not have a name property
              Type: FrameShape
              Size (mm): [40, 0]
            Shapes does not have a name property
              Type: FrameShape
              Size (mm): [61, 58]
            Shapes does not have a name property
              Type: FrameShape
              Size (mm): [91, 86]

A user who looked at ``build.odt`` for themselves might say that it contains three images not the two reported by |extract_graphics_py|_.
Why the discrepancy? |extract_graphics_py|_ only saves linked graphics, and only two were added by :py:meth:`.Write.add_image_link`.
The other image was inserted using :py:meth:`.Write.add_image_shape` which creates an image shape.

The number of shapes reported by |extract_graphics_py|_ may also confuse the user – why are there three rather than two?
The only shapes added to the document were an image and a line.

The names of the services gives a clue: the second and third shapes are the expected GraphicObjectShape_ and LineShape_, but some are
text frame (``FrameShape``) added by :py:meth:`.Write.add_text_frame`. Although this frame is an instance of the TextFrame_ service, it's reported as a ``FrameShape``.
That's a bit mysterious because there's no ``FrameShape`` service in the Office documentation.

.. _ch08_find_save_graphic:

8.3.1 Finding and Saving Text Graphics in a Document
----------------------------------------------------

:py:meth:`.Write.get_text_graphics` returns a list of XGraphic_ objects:
first it retrieves a collection of the graphic links in the document, then iterates through them creating an XGraphic_ object for each one:


.. tabs::

    .. code-tab:: python

        # in Write class
        @classmethod
        def get_text_graphics(cls, text_doc: XTextDocument) -> List[XGraphic]:
            try:
                xname_access = cls.get_graphic_links(text_doc)
                if xname_access is None:
                    raise ValueError("Unable to get Graphic Links")
                names = xname_access.getElementNames()

                pics: List[XGraphic] = []
                for name in names:
                    graphic_link = None
                    try:
                        graphic_link = xname_access.getByName(name)
                    except UnoException:
                        pass
                    if graphic_link is None:
                        Lo.print(f"No graphic found for {name}")
                    else:
                        try:
                            xgraphic = ImagesLo.load_graphic_link(graphic_link)
                            if xgraphic is not None:
                                pics.append(xgraphic)
                        except Exception as e:
                            Lo.print(f"{name} could not be accessed:")
                            Lo.print(f"    {e}")
                return pics
            except Exception as e:
                raise Exception(f"Get text graphics failed:") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Graphic objects are accessed with XTextGraphicObjectsSupplier_, as implemented by :py:meth:`.Write.get_graphic_links`:

.. tabs::

    .. code-tab:: python

        # in Write class
        @staticmethod
        def get_graphic_links(doc: XComponent) -> XNameAccess | None:
            ims_supplier = Lo.qi(XTextGraphicObjectsSupplier, doc, True)

            xname_access = ims_supplier.getGraphicObjects()
            if xname_access is None:
                Lo.print("Name access to graphics not possible")
                return None

            if not xname_access.hasElements():
                Lo.print("No graphics elements found")
                return None

            return xname_access

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Back in :py:meth:`.Write.get_text_graphics`, each graphic is loaded by calling :py:meth:`.ImagesLo.load_graphic_link`. It loads an image from the URL associated with a link:

.. tabs::

    .. code-tab:: python

        # in ImagesLo class
        @staticmethod
        def load_graphic_link(graphic_link: object) -> XGraphic:
            xprops = Lo.qi(XPropertySet, graphic_link, True)

            try:
                graphic = xprops.getPropertyValue("Graphic")
                if graphic is None:
                    raise Exception("Grapich is None")
                return graphic
            except Exception as e:
                raise Exception(f"Unable to retrieve graphic") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Note that the XGraphic **ARE** extracted from the document instead of loaded from a URL.
:py:meth:`.ImagesLo.load_graphic_file` can be used to load a graphic from a file.

.. note::

    See Tomaz's development blog `Part 1 <https://tomazvajngerl.blogspot.com/2018/01/improving-image-handling-in-libreoffice.html>`_ and `Part 2 <https://tomazvajngerl.blogspot.com/2018/03/improving-image-handling-in-libreoffice.html>`_
    for more information on why "GraphicURL" is no longer recommended.

    And `GraphicURL no longer works in 6.1.0.3 <https://ask.libreoffice.org/t/graphicurl-no-longer-works-in-6-1-0-3/35459>`_

Back in |extract_graphics|_, the XGraphic_ objects are saved as PNG files, and their pixel sizes reported:

.. tabs::

    .. code-tab:: python

        text_doc = WriteDoc(Write.open_doc(fnm=self._fnm, loader=loader))
        pics = text_doc.get_text_graphics()
        print(f"Num. of text graphics: {len(pics)}")

        # save text graphics to files
        for i, pic in enumerate(pics):
            img_file = self._out_dir / f"graphics{i}.png"
            ImagesLo.save_graphic(pic=pic, fnm=img_file)
            sz = cast(Size, Props.get(pic, "SizePixel"))
            print(f"Image size in pixels: {sz.Width} X {sz.Height}")
        print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.ImagesLo.save_graphic` utilizes the graphic provider's ``XGraphicProvider.storeGraphic()`` method:


.. tabs::

    .. code-tab:: python

        gprovider.storeGraphic(pic, png_props)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Its second argument is an array of PropertyValue_ objects, not a PropertySet_.
:py:class:`~.props.Props` utility class provides several functions for creating PropertyValue_ instances,
which are a variant of the ``{name=value}`` pair idea, but with extra data fields. One such function is:

.. tabs::

    .. code-tab:: python

        png_props = Props.make_props(URL=FileIO.fnm_to_url(fnm), MimeType=f"image/{im_format}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

It's passed an array of names and values, which are paired up as PropertyValue_ objects, and returned in a tuple of PropertyValue_.

In :py:meth:`~.ImagesLo.save_graphic`, these methods are used like so:

.. tabs::

    .. code-tab:: python

        gprovider = Lo.create_instance_mcf(XGraphicProvider, "com.sun.star.graphic.GraphicProvider")

        # set up properties for storing the graphic
        png_props = Props.make_props(URL=FileIO.fnm_to_url(fnm), MimeType=f"image/{im_format}")

        gprovider.storeGraphic(pic, png_props)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The idea is to call ``XGraphicProvider.saveGraphics()`` with the ``URL`` and ``MimeType`` properties set – the URL is for the image file, and the ``mimetype`` is an image type (e.g. ``image/png``).

:py:meth:`~.ImagesLo.save_graphic` is coded as:

.. tabs::

    .. code-tab:: python

        # in ImagesLo class
        @staticmethod
        def save_graphic(pic: XGraphic, fnm: PathOrStr, im_format: str) -> None:
            print(f"Saving graphic in '{fnm}'")

            if pic is None:
                print("Supplied image is null")
                return

            gprovider = Lo.create_instance_mcf(XGraphicProvider, "com.sun.star.graphic.GraphicProvider")
            if gprovider is None:
                print("Graphic Provider could not be found")
                return

            png_props = Props.make_props(URL=FileIO.fnm_to_url(fnm), MimeType=f"image/{im_format}")

            try:
                gprovider.storeGraphic(pic, png_props)
            except Exception as e:
                print("Unable to save graphic")
                print(f"    {e}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Other possible image MIME types include ``gif``, ``jpeg``, ``wmf``, and ``bmp``.
For instance, this call will save the image as a GIF: ``ImagesLO.save_graphic(pic, f"graphics{i}.gif", "gif")``
The printed output from :py:meth:`~.ImagesLo.save_graphic` contains another surprise:

.. code-block:: text

    Num. of text graphics: 2
    Saving graphic in graphics1.png
    Image size in pixels: 319 x 274
    Saving graphic in graphics2.png
    Image size in pixels: 319 x 274

The two saved graphics are the same size, but the second image is bigger inside the document.
The discrepancy is because the rendering of the image in the document is bigger, scaled up to fit the enclosing frame; the original image is unchanged.

.. _ch08_find_shapes:

8.3.2 Finding the Shapes in a Document
--------------------------------------

The report on shape block of code in |extract_graphics_py|_ reports on the shapes found in the document.
The relevant code fragment is:

.. tabs::

    .. code-tab:: python

        # code fragment from extract_graphics.py
        # report on shapes in the doc
        draw_page = text_doc.get_draw_page()
        shapes = Draw.get_shapes(draw_page.component)
        if shapes:
            print()
            print(f"No. of draw shapes: {len(shapes)}")

            for shape in shapes:
                Draw.report_pos_size(shape)
            print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Shapes are accessed with the ``XDrawPageSupplier.getDrawPage()`` method, which returns a single XDrawPage_:

.. tabs::

    .. code-tab:: python

        # in Write class
        @staticmethod
        def get_shapes(text_doc: XTextDocument) -> XDrawPage:
            draw_page_supplier = mLo.Lo.qi(XDrawPageSupplier, text_doc, True)
            return draw_page_supplier.getDrawPage()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

XDrawPage_'s usual role is to represent the canvas in Office's Draw, or a slide in Impress, and so plays an important role in :ref:`part03`.
Several support functions inside that part's :py:class:`~.draw.Draw` class will be used here.

``XDrawPageSupplier.getDrawPage()`` returns a single XDrawPage_ for the entire text document.
That doesn't mean that the shapes all have to occur on a single text page,
but rather that all the shapes spread across multiple text pages are collected into a single draw page.

XDrawPage_ inherits from XShapes_ and XindexAccess_, as shown in :numref:`ch08fig_xdrawpage_inherit`, which means that a page can be viewed as a indexed collection of shapes.

..
    figure 4

.. cssclass:: diagram invert

    .. _ch08fig_xdrawpage_inherit:
    .. figure:: https://user-images.githubusercontent.com/4193389/202319688-744c27e5-5d56-457a-91c0-d5c86466cf71.png
        :alt: Partial Inheritance Hierarchy for XDrawPage
        :figclass: align-center

        :Partial Inheritance Hierarchy for XDrawPage_

:py:meth:`.Draw.get_shapes` uses this idea to iterate through the draw page and store the shapes in a list:

.. tabs::

    .. code-tab:: python

        # in Draw class (overload method, simplified)
        @classmethod
        def get_shapes(cls, slide: XDrawPage) -> List[XShape]:
            if slide.getCount() == 0:
                Lo.print("Slide does not contain any shapes")
                return []

            shapes: List[XShape] = []
            for i in range(slide.getCount()):
                try:
                    shapes.append(Lo.qi(XShape, slide.getByIndex(i), True))
                except Exception as e:
                    continue
            return shapes

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        - :odev_src_draw_meth:`get_shapes`

XShape_ is part of the |d_shape|_ service, which contains many shape-related properties.

XShape_ inherits XShapeDescriptor_, which includes a ``getShapeType()`` method for returning the shape type as a string.
:numref:`ch08fig_shape_xshape` summarizes these details.

..
    figure 5

.. cssclass:: diagram invert

    .. _ch08fig_shape_xshape:
    .. figure:: https://user-images.githubusercontent.com/4193389/202321353-12ebe181-09ad-40cb-b78c-5f2811ed779e.png
        :alt: The Shape Service and XShape Interface
        :figclass: align-center

        :The |d_shape|_ Service and XShape_ Interface

:py:meth:`.Draw.show_shape_info` accesses the |d_shape|_ service associated with an XShape_ reference, and prints its ``XOrder`` property.
This number indicates the order that the shapes were added to the document.

.. tabs::

    .. code-tab:: python

        # in Draw class
        @classmethod
        def show_shape_info(cls, shape: XShape) -> None:
            print(f"  Shape service: {shape.getShapeType()}; z-order: {cls.get_zorder(shape)}")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`~.Draw.show_shape_info` also calls the inherited ``XShapeDescriptor.getShapeType()`` method to report the shape's service name.

.. _ch08_alt_access_shape:

8.3.3 Another Way of Accessing Drawing Shapes
---------------------------------------------

The XDrawPageSupplier_ documentation states that this interface is deprecated, although what's meant to replace it isn't clear.
It would seem that is XTextShapesSupplier_, although it did not supply anything.
For example, the following always reports that the supplier is ``None``:

.. tabs::

    .. code-tab:: python

        # this supplier is not created; Lo.qi() returns None
        shps_supp = Lo.qi(XTextShapesSupplier, text_doc)
        if shps_supp is None:
            print("Could not obtain text shapes supplier")
        else:
            print(f"No. of text shapes: {shps_supp.getShapes().getCount()}")
    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. |build_doc| replace:: Build Doc
.. _build_doc: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_doc

.. |extract_graphics| replace:: Extract Graphics
.. _extract_graphics: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_extract_graphics

.. |extract_graphics_py| replace:: extract_graphics.py
.. _extract_graphics_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/writer/odev_extract_graphics/extract_graphics.py

.. |d_shape| replace:: com.sun.star.drawing.Shape
.. _d_shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Shape.html

.. _BitmapTable: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1BitmapTable.html
.. _com.sun.star.drawing.Shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Shape.html
.. _com.sun.star.text.Shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1Shape.html
.. _GraphicObjectShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicObjectShape.html
.. _LineShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineShape.html
.. _PropertySet: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1beans_1_1PropertySet.html
.. _PropertyValue: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1beans_1_1PropertyValue.html
.. _TextFrame: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextFrame.html
.. _TextGraphicObject: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextGraphicObject.html
.. _XBitmap: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XBitmap.html
.. _XDrawPage: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPage.html
.. _XDrawPageSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPageSupplier.html
.. _XGraphic: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1graphic_1_1XGraphic.html
.. _XindexAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexAccess.html
.. _XNameContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameContainer.html
.. _XPropertySet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertySet.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
.. _XShapes: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShapes.html
.. _XShapeDescriptor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShapeDescriptor.html
.. _XTextContent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextContent.html
.. _XTextGraphicObjectsSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextGraphicObjectsSupplier.html
.. _XTextShapesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextShapesSupplier.html
