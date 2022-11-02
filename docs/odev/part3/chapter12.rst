.. _ch12:

*********************************************
Chapter 12. Examining a Draw/Impress Document
*********************************************

.. topic:: Overview

    Examining Slides/Pages; Page Layers; Styles

This chapter describes the |slide_info|_ example or more specificly |slide_info_py|_, which shows the basics of how to open and display a Draw or Impress file.
The code also illustrates how the slides or pages can be examined, including how to retrieve information on their layers and styles.

The |slide_info_py|_ main function:

.. tabs::

    .. code-tab:: python

        class SlidesInfo:
            def __init__(self, fnm: PathOrStr) -> None:
                FileIO.is_exist_file(fnm=fnm, raise_err=True)
                self._fnm = FileIO.get_absolute_path(fnm)

            def main(self) -> None:
                loader = Lo.load_office(Lo.ConnectPipe())

                try:
                    doc = Lo.open_doc(fnm=self._fnm, loader=loader)

                    if not Draw.is_shapes_based(doc):
                        Lo.print("-- not a drawing or slides presentation")
                        Lo.close_doc(doc)
                        Lo.close_office()
                        return

                    GUI.set_visible(is_visible=True, odoc=doc)
                    Lo.delay(1_000)  # need delay or zoom nay not occur

                    GUI.zoom(view=GUI.ZoomEnum.ENTIRE_PAGE)

                    print()
                    print(f"No. of slides: {Draw.get_slides_count(doc)}")
                    print()

                    # Access the first page
                    slide = Draw.get_slide(doc=doc, idx=0)

                    slide_size = Draw.get_slide_size(slide)
                    print(f"Size of slide (mm)({slide_size.Width}, {slide_size.Height})")
                    print()

                    self._report_layers(doc)
                    self._report_styles(doc)

                    Lo.delay(1_000)
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

:py:meth:`.Lo.open_doc` is capable of opening any Office document, and importing documents in other formats, so it's worthwhile checking the resulting
XComponent_ object before progressing. :py:meth:`.Draw.is_shapes_based` returns true if the file holds a Draw or Impress document:

.. tabs::

    .. code-tab:: python

        @staticmethod
        def is_shapes_based(doc: XComponent) -> bool:
            return mInfo.Info.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.DRAW) or mInfo.Info.is_doc_type(
                obj=doc, doc_type=mLo.Lo.Service.IMPRESS
            )

The document is made visible on-screen by calling :py:meth:`.GUI.set_visible`, and the application view is resized so all the drawing (or slide) is visible inside the window.
:py:meth:`.GUI.zoom` can be passed different :py:class:`.GUI.ZoomEnum` values for showing, ``ZoomEnum.PAGE_WIDTH``, the entire width of the page,
``ZoomEnum.ENTIRE_PAGE``, the entire page, ``ZoomEnum.OPTIMAL``. an optimal view that zooms in so all the 'data' on the page is visible without the empty space around it.
Alternatively, ``ZoomEnum.BY_VALUE`` with an interger value allows the user to supply a zooming percentage.
:abbreviation:`eg:` ``Draw.zoom(view=GUI.ZoomEnum.BY_VALUE, value=75)``

These two methods are defined using :py:meth:`.Lo.dispatch_cmd`, which was introduced at the end of :ref:`ch04`.

The call to :py:meth:`.Lo.delay` at the end of the zoom methods gives Office time to carry out the zooming before my code does anything else.
The same trick is utilized in the ``main()`` method, after the call to :py:meth:`.GUI.set_visible`.

.. seealso::

    .. cssclass:: mono

        - :odev_src_gui_meth:`zoom`:octicon:`code-square;1em;sd-text-info`

    - `Development/DispatchCommands <https://wiki.documentfoundation.org/Development/DispatchCommands>`_.

12.1 Accessing Slides/Pages
===========================

Most Draw class method names include the word 'slide' :abbreviation:`eg:` ( :py:meth:`.Draw.get_slides_count`, :py:meth:`.Draw.get_slide`, :py:meth:`.Draw.get_slide_size` ).
That's a bit misleading since most of them will work just as well with Draw or Impress documents.
For example, :py:meth:`.Draw.get_slides_count` will return 1 when applied to a newly created Draw document.

:py:meth:`.Draw.get_slides_count` calls :py:meth:`.Draw.get_slides` which returns an XDrawPages_ object; which supports a ``getCount()`` method:

.. tabs::

    .. code-tab:: python

        # in the Draw class (simplified)
        @classmethod
        def get_slides_count(cls, doc: XComponent) -> int:
            slides = cls.get_slides(doc)
            if slides is None:
                return 0
            return slides.getCount()

        @staticmethod
        def get_slides(doc: XComponent) -> XDrawPages:
            try:
                supplier = Lo.qi(XDrawPagesSupplier, doc, True)
                pages = supplier.getDrawPages()
                if pages is None:
                    raise DrawPageMissingError("Draw page supplier returned no pages")
                return pages
            except DrawPageMissingError:
                raise
            except Exception as e:
                raise DrawPageError("Error getting slides") from e

:py:meth:`~.Draw.get_slides` employs the XDrawPagesSupplier_ interface which is part of GenericDrawingDocument_ shown in :numref:`ch11fig_draw_and_presentation_services`.

:py:meth:`.Draw.get_slide` (note: no "s") treats the XDrawPages_ object as an indexed container of XDrawPage_ objects:

.. tabs::

    .. code-tab:: python

        # from draw class (simplified)
        @classmethod
        def get_slide(cls, doc: XComponent, idx: int) -> XDrawPage:
            # call: get_slide(cls, slides: XDrawPages, idx: int)
            return cls.get_slide(cls.get_slides(doc), idx)

        @classmethod
        def get_slide(cls, slides: XDrawPages, idx: int) -> XDrawPage:
            try:
                slide = Lo.qi(XDrawPage, slides.getByIndex(idx), True)
                return slide
            except IndexOutOfBoundsException:
                raise IndexError(f"Index out of bounds: {idx}")
            except Exception as e:
                raise DrawError(f"Could not get slide: {idx}") from e

:py:meth:`.Draw.get_slide_size` returns a |awt_size|_ object created from looking up the ``Width`` and ``Height`` properties of the supplied slide/page:

.. tabs::

    .. code-tab:: python

        # from Draw class (simplified)
        @staticmethod
        def get_slide_size(slide: XDrawPage) -> Size:
            try:
                props = Lo.qi(XPropertySet, slide)
                if props is None:
                    raise PropertySetMissingError("No slide properties found")
                width = int(props.getPropertyValue("Width"))
                height = int(props.getPropertyValue("Height"))
                return Size(round(width / 100), round(height / 100))
            except Exception as e:
                raise SizeError("Could not get shape size") from e

These ``Width`` and ``Height`` properties are stored in XDrawPage_'s GenericDrawPage_ service, shown in :numref:`ch11fig_some_drawpage_services`.

.. important::

    The :py:class:`~.draw.Draw` class specifies measurements in millimeters rather than Office's 1/100 mm units.
    For instance, :py:meth:`.Draw.get_slide_size` would return Office page dimensions of 28000 by 21000 as (280, 210).

.. seealso::

    .. cssclass:: mono

        - :odev_src_draw_meth:`get_slide`:octicon:`code-square;1em;sd-text-info`
        - :odev_src_draw_meth:`get_slides`:octicon:`code-square;1em;sd-text-info`
        - :odev_src_draw_meth:`get_slides_count`:octicon:`code-square;1em;sd-text-info`
        - :odev_src_draw_meth:`get_slide_size`:octicon:`code-square;1em;sd-text-info`

12.2 Page Layers
================

A Draw or Impress page consists of five layers called ``layout``, ``controls``, ``measurelines``, ``background``, and ``backgroundobjects``.
The first three are described in the Draw user guide, but ``measurelines`` is called "Dimension Lines".

Probably ``layout`` is the most important layer since that's where shapes are located.
Form controls (e.g. buttons) are added to "controls", which is always the top-most layer.
``background``, and ``backgroundobjects`` refer to the master page graphic and any shapes on that page.

Each layer can be made visible or invisible independent of the others. It's also possible to create new layers.

``_report_layers()`` in |slide_info_py|_ prints each layer's properties:

.. tabs::

    .. code-tab:: python

        # in slide_info.py
        def _report_layers(self, doc: XComponent) -> None:
            lm = Draw.get_layer_manager(doc)
            for i in range(lm.getCount()):
                try:
                    Props.show_obj_props(f" Layer {i}", lm.getByIndex(i))
                except:
                    pass
            layer = Draw.get_layer(doc=doc, layer_name=DrawingLayerKind.BACK_GROUND_OBJECTS)
            Props.show_obj_props("Background Object Props", layer)

:py:meth:`.Draw.get_layer_manager` obtains an XLayerManager_ instance which can be treated as an indexed container of XLayer_ objects.
:py:meth:`.Draw.get_layer` converts the XLayerManager_ into a named container, so it can be searched by layer name.

Typical output from ``_report_layers()`` is:

.. code::

    Layer 0 Properties
      Description: 
      IsLocked: False
      IsPrintable: True
      IsVisible: True
      Name: layout
      Title: 

    Layer 1 Properties
      Description: 
      IsLocked: False
      IsPrintable: True
      IsVisible: True
      Name: background
      Title: 

    Layer 2 Properties
      Description: 
      IsLocked: False
      IsPrintable: True
      IsVisible: True
      Name: backgroundobjects
      Title: 

    Layer 3 Properties
      Description: 
      IsLocked: False
      IsPrintable: True
      IsVisible: True
      Name: controls
      Title: 

    Layer 4 Properties
      Description: 
      IsLocked: False
      IsPrintable: True
      IsVisible: True
      Name: measurelines
      Title: 

    Background Object Props Properties
      Description: 
      IsLocked: False
      IsPrintable: True
      IsVisible: True
      Name: backgroundobjects
      Title: 

Each layer contains six properties. Four are defined by the Layer service; use`` lodoc layer service drawing`` to see its documentation.
The most useful property is probably ``IsVisible`` which toggles the layer's visibility.

12.3 Styles
===========

Draw and Impress utilize the same style organization as text documents, which was explained in :ref:`ch06`. :numref:`ch12fig_draw_impress_style_and_props` shows its structure.

..
    Figure 1

.. cssclass:: diagram invert

    .. _ch12fig_draw_impress_style_and_props:
    .. figure:: https://user-images.githubusercontent.com/4193389/199369511-8ac7e2d3-6d75-40b0-ab5f-5d131dc99c96.png
        :alt: Draw/Impress Style Families and their Property Sets
        :figclass: align-center

        :Draw/Impress Style Families and their Property Sets.

The style family names are different from those in text documents. The ``Default`` style family corresponds to the styles defined in a document's default master page.

:numref:`ch12fig_impress_default_master_pg` shows this master page in Impress.

..
    Figure 2

.. cssclass:: screen_shot invert

    .. _ch12fig_impress_default_master_pg:
    .. figure:: https://user-images.githubusercontent.com/4193389/199370492-7f386e5f-079c-4992-b11d-66f4a6552657.png
        :alt:  The Default Master Page in Impress.
        :figclass: align-center

        :The Default Master Page in Impress.

The master page (also known as a template in Impress' GUI) contains style information related to the title,
seven outline levels and background areas (e.g. the date, the footer, and the slide number in :numref:`ch12fig_impress_default_master_pg`).
Not all the master page styles are shown in :numref:`ch12fig_impress_default_master_pg`; for instance, there's a subtitle style, notes area, and a header.

If a slide deck is formatted using a master page (Impress template) other than ``Default``, such as ``Inspiration``,
then the style family name will be changed accordingly. The ``Inspiration`` family contains the same properties (styles) as ``Default``, but with different values.

.. todo::

    Chapte 12, Add link to chapter 17

Details on coding with master pages and Impress templates are given in the |master_use|_ and |points_builder|_ examples in Chapter 17.

.. todo::

    Chapte 12, Add link to chapter 14

The other Draw/Impress style families are ``cell``, ``graphics`` and ``table``. ``table`` and ``cell`` contain styles which affect the colors used to draw a table and its cells.
``graphics`` affects the appearance of shapes. Examples of how to use the ``graphics`` style family are given in the |draw_picture|_ example in Chapter 14.

The ``_report_styles()`` method inside |slide_info_py|_ is:

.. tabs::

    .. code-tab:: python

        def _report_styles(self, doc: XComponent) -> None:
            style_names = Info.get_style_family_names(doc)
            print("Style Families in this document:")
            Lo.print_names(style_names)
            # usually: "Default"  "cell"  "graphics"  "table"
            # Default is the name of the default Master Page template inside Office

            for name in style_names:
                con_names = Info.get_style_names(doc=doc, family_style_name=name)
                print(f'Styles in the "{name}" style family:')
                Lo.print_names(con_names)

The method prints the names of the style families, and the names of the styles (property sets) inside each family. Typical output is:

.. code::

    Style Families in this document:
    No. of names: 4
      'cell'  'Default'  'graphics'  'table'

    Styles in the "Default" style family:
    No. of names: 14
      'background'  'backgroundobjects'  'notes'  'outline1'
      'outline2'  'outline3'  'outline4'  'outline5'
      'outline6'  'outline7'  'outline8'  'outline9'
      'subtitle'  'title'

    Styles in the "cell" style family:
    No. of names: 34
      'blue1'  'blue2'  'blue3'  'bw1'
      'bw2'  'bw3'  'default'  'earth1'
      'earth2'  'earth3'  'gray1'  'gray2'
      'gray3'  'green1'  'green2'  'green3'
      'lightblue1'  'lightblue2'  'lightblue3'  'orange1'
      'orange2'  'orange3'  'seetang1'  'seetang2'
      'seetang3'  'sun1'  'sun2'  'sun3'
      'turquoise1'  'turquoise2'  'turquoise3'  'yellow1'
      'yellow2'  'yellow3'

    Styles in the "graphics" style family:
    No. of names: 40
      'A4'  'A4'  'Arrow Dashed'  'Arrow Line'
      'Filled'  'Filled Blue'  'Filled Green'  'Filled Red'
      'Filled Yellow'  'Graphic'  'Heading A0'  'Heading A4'
      'headline'  'headline1'  'headline2'  'Lines'
      'measure'  'Object with no fill and no line'  'objectwitharrow'  'objectwithoutfill'
      'objectwithshadow'  'Outlined'  'Outlined Blue'  'Outlined Green'
      'Outlined Red'  'Outlined Yellow'  'Shapes'  'standard'
      'Text'  'text'  'Text A0'  'Text A4'
      'textbody'  'textbodyindent'  'textbodyjustfied'  'title'
      'Title A0'  'Title A4'  'title1'  'title2'

    Styles in the "table" style family:
    No. of names: 11
      'blue'  'bw'  'default'  'earth'
      'green'  'lightblue'  'orange'  'seetang'
      'sun'  'turquoise'  'yellow'

.. |awt_size| replace:: com.sun.star.awt.Size
.. _awt_size: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1awt_1_1Size.html

.. |slide_info| replace:: Slides Info
.. _slide_info: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_slides_info

.. |slide_info_py| replace:: slide_info.py
.. _slide_info_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/impress/odev_slides_info/slides_info.py

.. |master_use| replace:: master use
.. _master_use: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_master_use

.. |points_builder| replace:: points builder
.. _points_builder: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_points_builder

.. |draw_picture| replace:: draw picture
.. _draw_picture: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_draw_picture

.. _GenericDrawingDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GenericDrawingDocument.html
.. _GenericDrawPage: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GenericDrawPage.html
.. _XComponent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XComponent.html
.. _XDrawPage: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPage.html
.. _XDrawPages: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPages.html
.. _XDrawPagesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPagesSupplier.html
.. _XLayer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XLayer.html
.. _XLayerManager: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XLayerManager.html