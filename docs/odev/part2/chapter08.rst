.. _ch08:

**************************
Chapter 8. Graphic Content
**************************

.. topic:: Overview

    Graphics; Linked Images/Shapes

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

8.1 Linking a Graphic Object to a Document
==========================================

Adding an image to a text document follows the same steps as other text content, as shown in :py:meth:`.Write.add_image_link`:

.. tabs::

    .. code-tab:: python

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
            _Events().trigger(WriteNamedEvent.IMAGE_LNIK_ADDING, cargs)
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
            _Events().trigger(WriteNamedEvent.IMAGE_LNIK_ADDED, EventArgs.from_args(cargs))
            return True

The TextGraphicObject_ service doesn't offer a ``XTextGraphicObject`` interface, so :py:meth:`.Lo.create_instance_msf`
returns an XTextContent_.

The interface is also converted to XPropertySet_ because several properties need to be set.
The frame is anchored, and the image's filename is assigned to ``GraphicURL`` (after being changed into a URL).

The image's size on the page depends on the dimensions of its enclosing frame, which are set in the "Width" and "Height" properties:

.. tabs::

    .. code-tab:: python

        props.setPropertyValue("Width", 4_500) # 45 mm width
        props.setPropertyValue("Height", 4_000) # 40 mm height

The values are in ``1/100`` mm units, so ``4500`` is ``45`` mm or ``4.5 cm``.

If these properties aren't explicitly set then the frame size defaults to being the width and height of the image.

In more realistic code, the width and height properties would be calculated as some scale factor of the image's size,
as measured in ``1/100 mm`` units not pixels.
These dimensions are available if the image file is loaded as an XGraphic_ object, as shown in :py:meth:`.ImagesLo.get_size_100mm`:

.. tabs::

    .. code-tab:: python

        @classmethod
        def get_size_100mm(cls, im_fnm: PathOrStr) -> Size:
            graphic = cls.load_graphic_file(im_fnm)
            return Props.get_property(prop_set=graphic, name="Size100thMM")

        @staticmethod
        def load_graphic_link(graphic_link: object) -> XGraphic:
            gprovider = Lo.create_instance_mcf(XGraphicProvider, "com.sun.star.graphic.GraphicProvider", raise_err=True)

            xprops = Lo.qi(XPropertySet, graphic_link, True)

            try:
                gprops = Props.make_props(URL=str(xprops.getPropertyValue("GraphicURL")))
                return gprovider.queryGraphic(gprops)
            except Exception as e:
                raise Exception(f"Unable to retrieve graphic") from e

Displaying the image at a scaled size is possible by combining :py:meth:`.ImagesLo.get_size_100mm` and :py:meth:`.Write.add_image_link`:

.. tabs::

    .. code-tab:: python

        img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)

        # enlarge by 1.5x
        h = round(img_size.Height * 1.5)
        w = round(img_size.Width * 1.5)
        Write.add_image_link(doc, cursor, im_fnm, w, h)

A possible drawback of :py:meth:`.Write.add_image_link` is that the document only contains a link to the image.
This becomes an issue if you save the document in a format other than ``.odt``.
In particular, when saved as a Word ``.doc`` file, the link is lost.

8.2 Adding a Graphic to a Document as a Shape
=============================================

.. todo::

    Chapter 8.2, add link to part 3.

An alternative to inserting a graphic as a link is to treat it as a shape.
Shapes will be discussed at length in Part 3, so I won't go into much detail about them here.
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

8.2.1 Creating an Image Shape
-----------------------------

The |build_doc|_ example adds an image shape to the document by calling :py:meth:`.Write.add_image_shape`:

.. tabs::

    .. code-tab:: python

        # code fragment from build doc
        # add image as shape to page
        append("Image as a shape: ")
        Write.add_image_shape(cursor=cursor, fnm=im_fnm)
        Write.end_paragraph(cursor)

:py:meth:`.Write.add_image_shape` comes in two versions: with and without width and height arguments.
A shape with no explicitly set width and height properties is rendered as a miniscule image (about 1 mm wide).
Call me old-fashioned, but I want to see the graphic, so :py:meth:`.Write.add_image_shape` calculates the picture's
dimensions if none are supplied by the user.

Another difference between image shape and image link is how the content's ``GraphicURL`` property is employed.
The image link version contains its URL, while the image shape's ``GraphicURL`` stores its bitmap as a string.

The code for :py:meth:`.Write.add_image_shape`:

.. tabs::

    .. code-tab:: python

        @classmethod
        def add_image_shape(cls, cursor: XTextCursor, fnm: PathOrStr, width: int = 0, height: int = 0) -> bool:
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
                gos = Lo.create_instance_msf(XTextContent, "com.sun.star.drawing.GraphicObjectShape", raise_err=True)

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

The image's size is calculated using :py:meth:`.ImagesLo.get_size_100mm` if the user doesn't supply a width and height,
and is used towards the end of the method.

An image shape is created using the GraphicObjectShape_ service, and its XTextContent_ interface is converted to XPropertySet_ for
assigning its properties, and to XShape_ for setting its size (see :numref:`ch08fig_shape_hierachy_parts`).
XShape_ includes a ``setSize()`` method.

8.2.2 Adding Other Graphics to the Document
===========================================

The graphic text content can be any sub-class of Shape. In the last section I created a GraphicObjectShape_ service, and accessed its XTextContent_ interface:

.. tabs::

    .. code-tab:: python

        gos = Lo.create_instance_msf(XTextContent, "com.sun.star.drawing.GraphicObjectShape")

In this section I'll use LineShape:

.. tabs::

    .. code-tab:: python

        ls = Lo.create_instance_msf(XTextContent, "com.sun.star.drawing.LineShape")

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

:py:meth:`~.Write.get_page_text_width` returns the writing width in ``1/100 mm`` units, which is scaled, then passed to :py:meth:`.Write.add_line_divider`:

.. tabs::

    .. code-tab:: python

        # code fragment in build doc
        text_width = Write.get_page_text_width(doc)
        # scale width by 0.5
        Write.add_line_divider(cursor=cursor, line_width=round(text_width * 0.5))

:py:meth:`~.Write.add_line_divider` creates a LineShape_ service with an XTextContent_ interface (see :numref:`ch08fig_shape_hierachy_parts`).
This is converted to XShape_ so its ``setSize()`` method can be passed the line width:

.. tabs::

    .. code-tab:: python

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
                cls.style_prev_paragraph(cursor=cursor, prop_val=ParagraphAdjust.CENTER, prop_name="ParaAdjust")

                cls.end_paragraph(cursor)
            except CreateInstanceMsfError:
                raise
            except MissingInterfaceError:
                raise
            except Exception as e:
                raise Exception("Insertion of graphic line failed") from e

The centering of the line is achieved by placing the shape in its own paragraph, then using :py:meth:`.Write.style_prev_paragraph` to center it.

8.3 Accessing Linked Images and Shapes
======================================

Work in Progress ...

.. |build_doc| replace:: Build Doc
.. _build_doc: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_build_doc

.. _BitmapTable: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1BitmapTable.html
.. _com.sun.star.drawing.Shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Shape.html
.. _com.sun.star.text.Shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1Shape.html
.. _GraphicObjectShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicObjectShape.html
.. _LineShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineShape.html
.. _TextGraphicObject: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextGraphicObject.html
.. _XBitmap: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XBitmap.html
.. _XGraphic: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1graphic_1_1XGraphic.html
.. _XNameContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XNameContainer.html
.. _XPropertySet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertySet.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
.. _XShapeDescriptor: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShapeDescriptor.html
.. _XTextContent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextContent.html