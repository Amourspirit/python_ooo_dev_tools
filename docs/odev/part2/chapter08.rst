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
returns an XTextContext_.

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

.. _TextGraphicObject: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextGraphicObject.html
.. _XTextContext: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextContent.html
.. _XPropertySet: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertySet.html
.. _XGraphic: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1graphic_1_1XGraphic.html