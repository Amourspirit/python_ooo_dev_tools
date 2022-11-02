.. _ch11:

**************************************
Chapter 11. Draw/Impress APIs Overview
**************************************

.. topic:: Overview

    Draw Pages and Master Pages; Draw Page Details; API Hierarchy Code Examples; Shapes in a Drawing; Shapes in a Presentation ; The Slide Show APIs

This part discusses the APIs for both Draw and Impress since the presentations API is an extension of Office's drawing functionality,
adding such things as slide-related shapes (e.g. for the title, subtitle, header, and footer), more data views (e.g. an handout mode), and slide shows.

You'll get a good feel for the APIs' capabilities by reading the Draw and Impress user guides,
downloadable from https://libreoffice.org/get-help/documentation.

Details about the APIs can be found in Chapter 9 of the Developer's Guide,
starting with |draw_pres_docs|_.
The guide can also be retrieved as a `PDF file <https://wiki.openoffice.org/w/images/d/d9/DevelopersGuide_OOo3.1.0.pdf>`_.

The guide's drawing and presentation examples are `online <https://api.libreoffice.org/examples/DevelopersGuide/examples.html#Drawing>`_,
and there's a short Draw example in `LibreOffice Java Examples <https://api.libreoffice.org/examples/examples.html#Java_examples>`_.

This chapter gives a broad overview of the drawing and presentation APIs, with some small code snippets to illustrate the ideas.
Subsequent chapters will return to these features in much more detail.

The APIs are organized around three services which subclass OfficeDocument, as depicted in :numref:`ch11fig_draw_and_presentation_services`.

..
    Figure 1

.. cssclass:: diagram invert

    .. _ch11fig_draw_and_presentation_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/199132193-3402c066-6abc-4308-afc8-a5ec04e77b98.png
        :alt: The Drawing and Presentation Document Services.
        :figclass: align-center

        :The Drawing and Presentation Document Services.

The DrawingDocument_ service, and most of its related services and interfaces are in Office's ``com.sun.star.drawing`` package (or module), which is documented as |draw_mod_api|_.

The presentation API is mostly located in Office's ``com.sun.star.presentation`` package (or module), which is documented as |pres_mod|_.

:numref:`ch11fig_draw_and_presentation_services_and_interfaces` shows a more detailed version of :numref:`ch11fig_draw_and_presentation_services`
which includes some of the interfaces defined by the services.

..
    Figure 2

.. cssclass:: diagram invert

    .. _ch11fig_draw_and_presentation_services_and_interfaces:
    .. figure:: https://user-images.githubusercontent.com/4193389/199133078-2ae26179-9b18-4741-9a7a-f404470e608c.png
        :alt: Drawing and Presentation Document Services and Interfaces.
        :figclass: align-center

        :Drawing and Presentation Document Services and Interfaces.

The interfaces highlighted in bold in :numref:`ch11fig_draw_and_presentation_services_and_interfaces` will be discussed in this chapter.

The DrawingDocument_ service is pretty much empty, with the real drawing 'action' in GenericDrawingDocument_ (which is in the ``com.sun.star.drawing`` package).
PresentationDocument_ subclasses GenericDrawingDocument_ to inherit its drawing capabilities,
and adds features for slide shows (via the XPresentationSupplier_ and XCustomPresentationSupplier_ interfaces).

The word "presentation" is a little overloaded in the API – PresentationDocument_ corresponds to the slide deck,
while ``XPresentationSupplier.getPresentation()`` returns an XPresentation_ object which represents a slide show.

11.1 Draw Pages and Master Pages
================================

A drawing (or presentation) document consists of a series of draw pages, one for each page (or slide) inside the document.
Perhaps the most surprising aspect of this is that a Draw document can contain multiple pages.

A document can also contain one or more master pages. A master page contains drawing/slide elements which appear on multiple draw page.
This idea is probably most familiar from slide presentations where a master page holds the header, footer, and graphics that appear on every slide.

As illustrated in  :numref:`ch11fig_draw_and_presentation_services`, GenericDrawingDocument_ supports an XDrawPagesSupplier_ interface whose ``getDrawPages()``
returns an XDrawPages_ object. It also has an XMasterPagesSupplier_ whose  ``getMasterPages()`` returns the master pages as an object.
Office views master pages as special kinds of draw pages, and so ``getMasterPages()`` also returns an XDrawPages_ object.

Note the "s" in XDrawPages_: an XDrawPages_ object is an indexed container of XDrawPage_ (no "s") objects, as illustrated by :numref:`ch11fig_xdrawpages_interface`.

..
    Figure 3

.. cssclass:: diagram invert

    .. _ch11fig_xdrawpages_interface:
    .. figure:: https://user-images.githubusercontent.com/4193389/199133801-568fc9f4-5f03-4ceb-a3f7-7b4bbd537e3b.png
        :alt: The X Draw Pages Interface
        :figclass: align-center

        :The XDrawPages_ Interface

Since XDrawPages_ inherit XIndexAccess_, its elements (pages) can be accessed using index-based lookup (i.e. to return page 0, page 1, etc.).

11.2 Draw Page Details
======================

A draw page is a collection of shapes: often text shapes, such as a title box or a box holding bulleted points.
But a shape can be many more things: an ellipse, a polygon, a bitmap, an embedded video, and so on.

This "page as shapes" notion is implemented by the API hierarchy shown in :numref:`ch11fig_drawpage_api_hierarchy`.

..
    Figure 4

.. cssclass:: diagram invert

    .. _ch11fig_drawpage_api_hierarchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/199135219-e593ea11-174c-4949-bd3a-11740ed23d74.png
        :alt: The API Hierarchy for a Draw Page
        :figclass: align-center

        :The API Hierarchy for a Draw Page.

XPresentationPage_ is the interface for a slide's page, but most of its functionality comes from XDrawPage_.
The XDrawPage_ interface doesn't do much either, except for inheriting XShapes_ (note the "s").
XShapes_ inherits XIndexAccess_, which means that an XShapes_ object can be manipulated as an indexed sequence of XShape_ objects.

The XDrawPage_ and XPresentationPage_ interfaces are supported by services, some of which are shown in :numref:`ch11fig_some_drawpage_services`.
These services are in some ways more important than the interfaces, since they contain many properties related to pages and slides.

..
    Figure 5

.. cssclass:: diagram invert

    .. _ch11fig_some_drawpage_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/199135727-ff5bc3e4-375f-42eb-9072-91818e6801d2.png
        :alt: Some of the Draw Page Services
        :figclass: align-center

        :Some of the Draw Page Services.

There are two ``DrawPage`` services in the Office API, one in the drawing package (DrawPage_), and another in the presentation package (Presentation_).
This is represented in :numref:`ch11fig_some_drawpage_services` by including the package names in brackets after the service names.
You can access the documentation for these services by typing ``lodoc drawpage service drawing`` and ``lodoc drawpage service presentation``.

No properties are defined in the drawing DrawPage_, instead everything is inherited from GenericDrawPage_.

There is "(??)" next to the XDrawPage_ and XPresentationPage_ interfaces in :numref:`ch11fig_some_drawpage_services` because they're not listed in the
GenericDrawPage_ and presentation DrawPage_ services in the documentation, but must be there because of the way that the code works.
Also, the documentation for GenericDrawPage_ lists XShapes_ as an interface, rather than XDrawPage_.

11.3 API Hierarchy Code Examples
================================

Some code snippets will help clarify the hierarchies shown in :numref:`ch11fig_draw_and_presentation_services_and_interfaces` - :numref:`ch11fig_some_drawpage_services`.
The following lines load a Draw (or Impress) document called "foo" as an XComponent_ object.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.lo import Lo

        loader = Lo.load_office(Lo.ConnectPipe())
        doc = Lo.open_doc(fnm="foo", loader=loader)

A common next step is to access the draw pages in the document using the XDrawPagesSupplier_ interface
shown in :numref:`ch11fig_draw_and_presentation_services_and_interfaces`:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.lo import Lo

        supplier = Lo.qi(XDrawPagesSupplier, doc)
        pages = supplier.getDrawPages() # XDrawPages

This code works whether the document is a sequence of draw pages (i.e. a Draw document) or
slides (i.e. an Impress document).

Using the ideas shown in :numref:`ch11fig_xdrawpages_interface`, a particular draw page is accessed based on
its index position. The first draw page in the document is retrieved with:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.lo import Lo

        page = Lo.qi(XDrawPage, pages.getByIndex(0))

Pages are numbered from 0, and a newly created document always contains one page.

The XDrawPage_ interface is supported by the GenericDrawPage service (see :numref:`ch11fig_some_drawpage_services`),
which holds the page's properties. The following snippet returns the width of the page and its page number:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.props import Props

        width =  int(Props.get(page, "Width"))
        page_number = int(Props.get(page, "Number"))

The "Width" and "Number" properties are documented in the GenericDrawPage_ service.

Once a single page has been retrieved, it's possible to access its shapes (as shown in :numref:`ch11fig_drawpage_api_hierarchy`).
The following code converts the XDrawPage_ object to XShapes_, and accesses the first XShape_ in its indexed container:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.lo import Lo

        shapes = Lo.qi(XShapes, page)
        shape = Lo.qi(XShape, shapes.getByIndex(0))

11.4 Shapes in a Drawing
========================

Shapes fall into two groups – drawing shapes that subclass the Shape service in ``com.sun.star.drawing``,
and presentation-related shapes which subclass the Shape service in ``com.sun.star.presentation``.
The first type are described here, and the presentation shapes in :ref:`ch011_shapes_in_a_presentation`.

:numref:`ch11fig_some_drawing_shapes` shows the ``com.sun.star.drawing`` Shape_ service and some of its subclasses.

..
    Figure 6

.. cssclass:: diagram invert

    .. _ch11fig_some_drawing_shapes:
    .. figure:: https://user-images.githubusercontent.com/4193389/199140156-6773e06f-cf89-4b9f-ba45-1c0d7bae5908.png
        :alt: Some of the Drawing Shapes
        :figclass: align-center

        :Some of the Drawing Shapes.

.. todo::

    | Chapte 11, Add link to chapter 13
    | Chapte 11, Add link to chapters 15

More about many of these shapes in Chapters 13 and 15, but you can probably guess what most of them do
– ``EllipseShape`` is for drawing ellipses and circles, ``LineShape`` is for lines and arrows, ``RectangleShape`` is for rectangles.

The two "??"s in :numref:`ch11fig_some_drawing_shapes` indicate that those services aren't shown in the online documentation, but appear in examples.

The hardest aspect of this hierarchy is searching it for information on a shape's properties.
Many general properties are located in the Shape_ service.
More specialized properties are located in the specific shape's service.
For instance, RectangleShape_ has a ``CornerRadius`` property which allows a rectangle's corners to be rounded to make it more button-like.

Unfortunately, most shapes inherit a lot more properties than just those in Shape. :numref:`ch11fig_rectangel_shape_props` shows a typical example
– RectangleShape_ inherits properties from at least eight services (I've not shown the complete hierarchy)!

..
    Figure 7

.. cssclass:: diagram invert

    .. _ch11fig_rectangel_shape_props:
    .. figure:: https://user-images.githubusercontent.com/4193389/199142104-0d81e264-c8ec-474d-8eeb-7707bfd192ca.png
        :alt: RectangleShape's Properties.
        :figclass: align-center

        :RectangleShape_'s Properties.

Aside from RectangleShape_ inheriting properties from Shape_, it also obtains its fill, shadow, line, and rotation attributes from the
FillProperties_, ShadowProperties_, LineProperties_, and RotationDescriptor_ services. For instance, to make the rectangle red, the ``FillColor`` property,
defined in the FillProperties_ service, must be set. The code for doing this is not complex:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.props import Props
        from ooodev.utils.color import CommonColor

        Props.set(shape, FillColor=CommonColor.RED)

The complication comes in knowing that a property called ``FillColor`` exists.
Follow RectangleShape_ link and look inside each inherited Property service until you find the relevant property.

.. tip::

    Thre is a `List of all members <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1RectangleShape-members.html>`_ link
    on the top right side of all API pages.

If the shape contains some text (e.g. the rectangle has a label inside it), and you want to change one of the text's properties,
then you'll need to look in the three property services above the Text service (see :numref:`ch11fig_rectangel_shape_props`).

Changing text requires that the text be selected first, which takes us back XText_ and :ref:`ch05`.
For example, the text height is changed to ``18pt`` by:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.props import Props
        from ooodev.utils.lo import Lo

        xtext = Lo.qi(XText, shape)
        cursor = xtext.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True) #  select all text
        Props.set(cursor, CharHeight=18);

First the shape is converted into an XText_ reference so that text selection can be done using a cursor.

The ``CharHeight`` property is inherited from the CharacterProperties_ service.

:numref:`ch11fig_rectangel_shape_props` doesn't show all the text property services.
For instance, there are also services called CharacterPropertiesComplex_ and ParagraphPropertiesComplex_.

.. _ch011_shapes_in_a_presentation:

11.5 Shapes in a Presentation
=============================

If the document is a slide deck, then presentation-related shapes will be subclasses of the Shape_ service.
Some of those shapes are shown in :numref:`ch11fig_some_presentation_shapes`.

..
    Figure 8

.. cssclass:: diagram invert

    .. _ch11fig_some_presentation_shapes:
    .. figure:: https://user-images.githubusercontent.com/4193389/199145787-e6ea86cf-01fc-485f-be5d-7dc48b27545c.png
        :alt: Some of the Presentation Shapes.
        :figclass: align-center

        :Some of the Presentation_ Shapes.

The |presentation_Shape|_ service doesn't subclass the  |drawing_Shape|_ service.
Instead, every presentation shape inherits the presentation Shape service and a drawing shape (usually TextShape_).
This means that all the presentation shapes can be treated as drawing shapes when being manipulated in code.

Most of the presentation shapes are special kinds of text shapes.
For instance, TitleTextShape_ and OutlinerShape_ are text shapes which usually appear automatically when you create a new slide inside
Impress – the slide's title is typed into the TitleTextShape_, and bulleted points added to OutlinerShape_. This is shown in  :numref:`ch11fig_two_presentation_shapes`.

..
    Figure 9

.. cssclass:: screen_shot invert

    .. _ch11fig_two_presentation_shapes:
    .. figure:: https://user-images.githubusercontent.com/4193389/199147338-03111ff1-6273-4a9c-9af4-c84317ec3e0b.png
        :alt: Two Presentation Shapes in a Slide.
        :figclass: align-center

        :Two Presentation Shapes in a Slide.

Using OutlinerShape_ as an example, its 'simplified' inheritance hierarchy looks like :numref:`ch11fig_outliner_hierarchy`.

..
    Figure 10

.. cssclass:: diagram invert

    .. _ch11fig_outliner_hierarchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/199147692-3701ce06-b468-416a-8917-bb20052e0615.png
        :alt: The Outliner Shape Hierarchy
        :figclass: align-center

        :The OutlinerShape_ Hierarchy.

OutlinerShape_ has at least nine property services that it inherits.

11.6 The Slide Show APIs
========================

One difference between slides and drawings is that the presentations API supports slide shows.
This extra functionality can be seen in :numref:`ch11fig_draw_and_presentation_services_and_interfaces` since the PresentationDocument_ service offers an XPresentationSupplier_ interface
which has a ``getPresentation()`` method for returning an XPresentation_ object. Don't be confused by the name – an XPresentation_ object represents a slide show.


XPresentation_ offers ``start()`` and ``end()`` methods for starting and ending a slide show,
and the Presentation_ service contains properties for affecting how the show progresses, as illustrated by :numref:`ch11fig_slide_show_presentation_services`.

..
    Figure 11

.. cssclass:: diagram invert

    .. _ch11fig_slide_show_presentation_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/199148818-713cac28-4045-48b4-b8a4-42b7fd74199b.png
        :alt: The Slide Show Presentation Services.
        :figclass: align-center

        :The Slide Show Presentation_ Services.

Code for starting a slide show for the "foo" document:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.lo import Lo

        loader = Lo.load_office(Lo.ConnectPipe())
        doc = Lo.open_doc("foo", loader)
        ps = Lo.qi(XPresentationSupplier, doc)
        Lo.qi(XPresentation, ps.getPresentation())
        show.start()

Alternatively a slide show can be started with a dispatch command.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.lo import Lo
        from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch

        loader = Lo.load_office(Lo.ConnectPipe())
        doc = Lo.open_doc("foo", loader)
        Lo.delay(500) #  wait half a sec
        Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)

The Presentation_ service is a bit lacking, so an extended service, Presentation2_, was added more recently.
It can access an XSlideShowController_ interface which gives finer-grained control over how the show progresses; examples will come later.


.. |draw_pres_docs| replace:: Drawing Documents and Presentation Documents
.. _draw_pres_docs: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Drawing_Documents_and_Presentation_Documents

.. |draw_mod_api| replace:: drawing Module
.. _draw_mod_api: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html

.. |pres_mod| replace:: presentation Module
.. _pres_mod: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1presentation.html

.. |drawing_Shape| replace:: com.sun.star.drawing.Shape
.. _drawing_Shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Shape.html

.. |presentation_Shape| replace:: com.sun.star.presentation.Shape
.. _presentation_Shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1Shape.html

.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _CharacterPropertiesComplex: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterPropertiesComplex.html
.. _DrawingDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1DrawingDocument.html
.. _DrawPage: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1DrawPage.html
.. _FillProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1FillProperties.html
.. _GenericDrawingDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GenericDrawingDocument.html
.. _GenericDrawPage: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GenericDrawPage.html
.. _LineProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineProperties.html
.. _OutlinerShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1OutlinerShape.html
.. _ParagraphPropertiesComplex: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphPropertiesComplex.html
.. _Presentation: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1Presentation.html
.. _Presentation: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1Presentation.html
.. _Presentation2: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1Presentation2.html
.. _PresentationDocument: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1PresentationDocument.html
.. _RectangleShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1RectangleShape.html
.. _RotationDescriptor: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1RotationDescriptor.html
.. _ShadowProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1ShadowProperties.html
.. _Shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Shape.html
.. _TextShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1TextShape.html
.. _TitleTextShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1TitleTextShape.html
.. _XComponent: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XComponent.html
.. _XCustomPresentationSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1presentation_1_1XCustomPresentationSupplier.html
.. _XDrawPage: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPage.html
.. _XDrawPages: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPages.html
.. _XDrawPagesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPagesSupplier.html
.. _XIndexAccess: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XIndexAccess.html
.. _XMasterPagesSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XMasterPagesSupplier.html
.. _XPresentation: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1presentation_1_1XPresentation.html
.. _XPresentationPage: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1presentation_1_1XPresentationPage.html
.. _XPresentationSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1presentation_1_1XPresentationSupplier.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
.. _XShapes: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShapes.html
.. _XSlideShowController: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1presentation_1_1XSlideShowController.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
