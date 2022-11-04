.. _ch13:

********************************
Chapter 13. Drawing Basic Shapes
********************************

.. topic:: Overview

    A Black Dashed Line; A Red Ellipse; Filled Rectangles; Text; Shape Names; A Transparent Circle and a Polar Line; A Math Formula as an OLE Shape; Polygons; Multiple Lines, Partial Elipses

This chapter contains an assortment of basic shape creation examples, including the following:

.. cssclass:: ul-list

    * simple shapes: line, ellipse, rectangle, text;
    * shape fills: solid, gradients, hatching, bitmaps;
    * an OLE shape (a math formulae);
    * polygons, multiple lines, partial ellipses.

The examples come from two files, |draw_picture|_ and |animate_bike|_. The ``show()`` function of |draw_picture_py|_:

.. tabs::

    .. code-tab:: python


        class DrawPicture:

            def show(self) -> None:
                loader = Lo.load_office(Lo.ConnectPipe())

                try:
                    doc = Draw.create_draw_doc(loader)
                    GUI.set_visible(is_visible=True, odoc=doc)
                    Lo.delay(1_000)  # need delay or zoom may not occur
                    GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                    curr_slide = Draw.get_slide(doc=doc, idx=0)
                    self._draw_shapes(curr_slide=curr_slide)

                    s = Draw.draw_formula(slide=curr_slide, formula="func e^{i %pi} + 1 = 0", x=70, y=20, width=75, height=40)
                    # Draw.report_pos_size(s)

                    self._anim_shapes(curr_slide=curr_slide)

                    s = Draw.find_shape_by_name(curr_slide, "text1")
                    Draw.report_pos_size(s)

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

|draw_picture_py|_ creates a new Draw document, and finishes by dispalying a :ref:`class_msg_box` shown in :numref:`ch13fig_msgbox_all_done` asking the user if they want to close the document.

.. tabs::

    .. code-tab:: python

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

.. cssclass:: screen_shot

    .. _ch13fig_msgbox_all_done:
    .. figure:: https://user-images.githubusercontent.com/4193389/199492083-75137d38-3bd4-4290-9972-5be7cf8e2d68.png
        :alt: Message Box - All Done
        :figclass: align-center

        :Message Box - All Done

:py:meth:`.Draw.create_draw_doc` is a wrapper around :py:meth:`.Lo.create_doc` to create a Draw document:

.. tabs::

    .. code-tab:: python

        # in the Draw class
        @staticmethod
        def create_draw_doc(loader: XComponentLoader) -> XComponent:
            return Lo.create_doc(doc_type=Lo.DocTypeStr.DRAW, loader=loader)

.. tabs::

    .. code-tab:: python

        # in the Draw class
        @staticmethod
        def create_impress_doc(loader: XComponentLoader) -> XComponent:
            return Lo.create_doc(doc_type=Lo.DocTypeStr.IMPRESS, loader=loader)

.. seealso::

    .. cssclass:: src-link

        - :odev_src_draw_meth:`create_draw_doc`
        - :odev_src_draw_meth:`create_impress_doc`

13.1 Drawing Shapes
===================

The ``_draw_shapes()`` method inside |draw_picture_py|_ draws the six shapes shown in :numref:`ch13fig_draw_shapes_six`.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch13fig_draw_shapes_six:
    .. figure:: https://user-images.githubusercontent.com/4193389/199504922-6029aa82-f986-45c6-8be3-2bd908e130a7.png
        :alt: The Six Shapes Drawn by draw Shapes.
        :figclass: align-center

        :The Six Shapes Drawn by ``_draw_shapes()``.

Almost every Draw method call :py:meth:`.Draw.make_shape` which creates a shape instance and sets its size and position on the page:

.. tabs::

    .. code-tab:: python

        # in the Draw class (simplified)
        @staticmethod
        def make_shape(shape_type: DrawingShapeKind | str, x: int, y: int, width: int, height: int) -> XShape:
            # parameters are in mm units
            try:
                shape = Lo.create_instance_msf(XShape, f"com.sun.star.drawing.{shape_type}", raise_err=True)
                shape.setPosition(Point(x * 100, y * 100))
                shape.setSize(Size(width * 100, height * 100))
                return shape
            except Exception as e:
                raise ShapeError(f'Unable to create shape "{shape_type}"') from e

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`make_shape`

The method assumes that the shape is defined inside the ``com.sun.star.drawing`` package, :abbreviation:`i.e.` that it's a shape which
subclasses |drawing_shape|_, like those in :numref:`ch11fig_some_drawing_shapes`.
The code converts the supplied (x, y) coordinate, width, and height from millimeters to Office's ``1/100 mm`` values.

The exact meaning of the position and the size of a shape is a little tricky.
If its width and height are positive, then the position is the top-left corner of the rectangle defined by those dimensions.
However, the user can supply negative dimensions, which means that "top-left corner" may be on the right or bottom of the rectangle
(see :numref:`ch13fig_office_store_shapes` (a)). Office handles this by storing the rectangle with a new top-left point,
so all the dimensions can be positive (see :numref:`ch13fig_office_store_shapes` (b)).

..
    figure 2

.. cssclass:: diagram invert

    .. _ch13fig_office_store_shapes:
    .. figure:: https://user-images.githubusercontent.com/4193389/199507795-c1de83cb-3754-4337-a8e4-2fa7a35811c4.png
        :alt: How Office Stores a Shape with a Negative Height.
        :figclass: align-center

        :How Office Stores a Shape with a Negative Height.

This means that your code should not assume that the position and size of a shape remain unchanged after being set with ``XShape.setPosition()`` and ``XShape.setSize()``.

:py:meth:`~.Draw.make_shape` is called by :py:meth:`.Draw.add_shape` which adds the generated shape to the page.
It also check if the (x, y) coordinate is located on the page. If it isn't, :py:meth:`.Draw.warn_position` prints a warning message.

.. tabs::

    .. code-tab:: python

        # in the Draw class (simplified)
        @classmethod
        def add_shape(
            cls, slide: XDrawPage, shape_type: DrawingShapeKind | str, x: int, y: int, width: int, height: int
        ) -> XShape:
            try:
                cls.warns_position(slide=slide, x=x, y=y)
                shape = cls.make_shape(shape_type=shape_type, x=x, y=y, width=width, height=height)
                slide.add(shape)
                return shape
            except ShapeError:
                raise
            except Exception as e:
                raise ShapeError("Error adding shape") from e

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`add_shape`

``_draw_shapes()`` in the |draw_picture_py|_ example is shown below. It creates the six shapes shown in  :numref:`ch13fig_draw_shapes_six`.

.. tabs::

    .. code-tab:: python

        def _draw_shapes(self, curr_slide: XDrawPage) -> None:
            line1 = Draw.draw_line(slide=curr_slide, x1=50, y1=50, x2=200, y2=200)
            Props.set(line1, LineColor=CommonColor.BLACK)
            Draw.set_dashed_line(shape=line1, is_dashed=True)

            # red ellipse; uses (x, y) width, height
            circle1 = Draw.draw_ellipse(slide=curr_slide, x=100, y=100, width=75, height=25)
            Props.set(circle1, FillColor=CommonColor.RED)

            # rectangle with different fills; uses (x, y) width, height
            rect1 = Draw.draw_rectangle(slide=curr_slide, x=70, y=100, width=75, height=25)
            Props.set(rect1, FillColor=CommonColor.LIME)

            text1 = Draw.draw_text(
                slide=curr_slide, msg="Hello LibreOffice", x=120, y=120, width=60, height=30, font_size=24
            )
            Props.set(text1, Name="text1")
            # Props.show_props("TextShape's Text Properties", Draw.get_text_properties(text1))

            # gray transparent circle; uses (x,y), radius
            circle2 = Draw.draw_circle(slide=curr_slide, x=40, y=150, radius=20)
            Props.set(circle2, FillColor=CommonColor.GRAY)
            Draw.set_transparency(shape=circle2, level=Intensity(25))

            # thick line; uses (x,y), angle clockwise from x-axis, length
            line2 = Draw.draw_polar_line(slide=curr_slide, x=60, y=200, degrees=45, distance=100)
            Props.set(line2, LineWidth=300)

There's a number of variations possible for each shape.
The following sections look at how the six shapes are drawn.

13.2 A Black Dashed Line
========================

:py:meth:`.Draw.draw_line` calls :py:meth:`.Draw.add_shape` to create a |drawing_line_shape|_ instance.
In common with other shapes, a line is defined in terms of its enclosing rectangle, represented by its top-left corner, width, and height.
:py:meth:`.Draw.draw_line` allows the programmer to define the line using its endpoints:

.. tabs::

    .. code-tab:: python

        # in the Draw class
        @classmethod
        def draw_line(cls, slide: XDrawPage, x1: int, y1: int, x2: int, y2: int) -> XShape:
            # make sure size is non-zero
            if (x1 == x2) and (y1 == y2):
                raise ValueError("Cannot create a line from a point")

            width = x2 - x1  # may be negative
            height = y2 - y1  # may be negative
            return cls.add_shape(
                slide=slide,
                shape_type=DrawingShapeKind.LINE_SHAPE,
                x=x1,
                y=y1,
                width=width,
                height=height,
            )

As mentioned above, Office will store a shape with a modified position and size if one or both of its dimensions is negative.
As an example, consider if :py:meth:`.Draw.draw_line` is called with the coordinates (10,20) and (20,10).
The call to :py:meth:`.Draw.add_shape` would be passed a positive width (``10mm``) and a negative height (``-10mm``).
This would be drawn as in :numref:`ch13fig_office_neg_shape` (a) but would be stored using the shape position and size in :numref:`ch13fig_office_neg_shape` (c).

..
    figure 3

.. cssclass:: diagram invert

    .. _ch13fig_office_neg_shape:
    .. figure:: https://user-images.githubusercontent.com/4193389/199515829-405bf789-9033-441d-9032-44e4ac5b2b9f.png
        :alt: How a Line with a Negative Height is Stored as a Shape
        :figclass: align-center

        :How a Line with a Negative Height is Stored as a Shape.

This kind of transformation may be important if your code modifies a shape after it has been added to the slide, as my animation examples do in the next chapter.

Back in |draw_picture_py|_'s ``_draw_shapes()``, the line's properties are adjusted.
The hardest part of this is finding the property's name in the API documentation, because properties are typically defined across multiple services,
including LineShape_, Shape_, FillProperties_, ShadowProperties_, LineProperties_, and RotationDescriptor_.
If the property is related to the shape's text then you should check TextProperties_, CharacterProperties_, and ParagraphProperties_ as well.
:numref:`ch11fig_rectangel_shape_props` shows the situation for RectangleShape, and its much the same for other shapes.
You should start looking in the documentation for the shape ( :abbreviation:`i.e.` use lodoc LineShape_ drawing ), and move up the hierarchy.

.. tip::

    Thre is a `List of all members <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineShape-members.html>`_ link
    on the top right side of all API pages.

You can click on the inheritance diagram at the top of the page ( :abbreviation:`e.g.` like the one in :numref:`ch13fig_line_shape_diagram` ) to look in the different services.

..
    figure 4

.. cssclass:: diagram invert

    .. _ch13fig_line_shape_diagram:
    .. figure:: https://user-images.githubusercontent.com/4193389/199562000-f5a1b03d-638b-4c2c-bebb-6ab026dd0d52.png
        :alt: The Line Shape Inheritance Diagram in the LibreOffice Online Documentation.
        :figclass: align-center

        :The LineShape_ Inheritance Diagram in the LibreOffice Online Documentation.

``_draw_shapes()`` will color the line black and make it dashed, which suggests that I should examine the LineProperties_ class.
Its relevant properties are ``LineColor`` for color and ``LineStyle`` and ``LineDash`` for creating dashes, as in :numref:`ch13fig_line_prop_rel`.

..
    figure 5

.. cssclass:: diagram invert

    .. _ch13fig_line_prop_rel:
    .. figure:: https://user-images.githubusercontent.com/4193389/199562708-410a32af-5b4b-4d73-a225-0f0f6ac4415f.png
        :alt: Relevant Properties in the Line Properties Class.
        :figclass: align-center

        :Relevant Properties in the LineProperties_ Class.

Line color can be set with a single call to :py:meth:`.Props.set`, but line dashing is a little more complicated, so is handled by :py:meth:`.Draw.set_dashed_line`:

.. tabs::

    .. code-tab:: python

        # in _draw_Shapes()
        Props.set(line1, LineColor=CommonColor.BLACK)
        Draw.set_dashed_line(shape=line1, is_dashed=True)

.. seealso::

    :ref:`module_color`

:py:meth:`.Draw.set_dashed_line` has to assign a LineStyle_ object to ``LineStyle`` and a LineDash_ object to ``LineDash``.
The line style is easily set since LineStyle_ is an enumeration with three possible values. A ``LineDash`` object requires more work:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)

        from ooo.dyn.drawing.line_dash import LineDash as LineDash
        from ooo.dyn.drawing.line_style import LineStyle as LineStyle

        @staticmethod
        def set_dashed_line(shape: XShape, is_dashed: bool) -> None:
            try:
                props = Lo.qi(XPropertySet, shape, True)
                if is_dashed:
                    ld = LineDash() #  create new struct
                    ld.Dots = 0
                    ld.DotLen = 100
                    ld.Dashes = 5
                    ld.DashLen = 200
                    ld.Distance = 200
                    props.setPropertyValue("LineStyle", LineStyle.DASH)
                    props.setPropertyValue("LineDash", ld)
                else:
                    # switch to solid line
                    props.setPropertyValue("LineStyle", LineStyle.SOLID)
            except Exception as e:
                raise ShapeError("Error setting dashed line property") from e

:py:meth:`~.Draw.set_dashed_line` can be used to toggle a line's dashes on or off.

.. note::

    :py:class:`~.draw.Draw` class import may enums and structures from ooouno_ package, including ``LineDash`` and ``LineStyle``.
    At runtime their values and constants are identical to ``uno's``. The advantage is there is a little magic taking place under the
    hood with ooouno_ imports in the ``dyn`` namespace. They behave like python objects without the ``uno`` limitations.

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`set_dashed_line`

13.3 A Red Ellipse
==================

A red ellipse is drawn using:

.. tabs::

    .. code-tab:: python

        # in _draw_Shapes()
        circle1 = Draw.draw_ellipse(slide=curr_slide, x=100, y=100, width=75, height=25)
        Props.set(circle1, FillColor=CommonColor.RED)

:py:meth:`.Draw.draw_ellipse` is similar to :py:meth:`.Draw.draw_line` except that an EllipseShape_ is created by :py:meth:`.Draw.add_shape`:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def draw_ellipse(cls, slide: XDrawPage, x: int, y: int, width: int, height: int) -> XShape:
            return cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.ELLIPSE_SHAPE, x=x, y=y, width=width, height=height
            )

The circle needs to be filled with a solid color, which suggests the setting of a property in FillProperties_.
A visit to the online documentation for EllipseShape_ reveals an inheritance diagram like the one in :numref:`ch13fig_ellipse_shape_inherit_diag`.

..
    figure 6

.. cssclass:: diagram

    .. _ch13fig_ellipse_shape_inherit_diag:
    .. figure:: https://user-images.githubusercontent.com/4193389/199569929-c6490409-98af-448a-9f69-8996aa282c43.png
        :alt: The Ellipse Shape Inheritance Diagram in the Libre Office Online Documentation.
        :figclass: align-center

        :The EllipseShape_ Inheritance Diagram in the LibreOffice Online Documentation.

Clicking on the FillProperties_ rectangle jumps to its documentation, which lists a ``FillColor`` property (see :numref:`ch13fig_fill_properties_rel_prop`).

..
    figure 7

.. cssclass:: diagram invert

    .. _ch13fig_fill_properties_rel_prop:
    .. figure:: https://user-images.githubusercontent.com/4193389/199571390-07a009dd-62e9-4cc2-baf8-29a714ef98a3.png
        :alt: Relevant Properties in the Fill Properties Class.
        :figclass: align-center

        :Relevant Properties in the FillProperties_ Class.

Both the ``FillColor`` and ``FillStyle`` properties should be set, but the default value for ``FillStyle`` is already ``FillStyle.SOLID``, which is what's needed.

13.4 A Rectangle with a Variety of Fills
========================================

The rectangle example in drawShapes() comes in six different colors:

.. tabs::

    .. code-tab:: python

        # in _draw_Shapes()
        # rectangle with different fills; uses (x, y) width, height
        rect1 = Draw.draw_rectangle(slide=curr_slide, x=70, y=100, width=75, height=25)
        Props.set(rect1, FillColor=CommonColor.LIME)
        

Work in progress ...

.. |animate_bike| replace:: Animate Bike
.. _animate_bike: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_animate_bike

.. |animate_bike_py| replace:: Animate Bike
.. _animate_bike_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/draw/odev_animate_bike/anim_bicycle.py

.. |draw_picture| replace:: Draw Picture
.. _draw_picture: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_draw_picture

.. |draw_picture_py| replace:: draw_picture.py
.. _draw_picture_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_draw_picture/draw_picture.py

.. |drawing_shape| replace:: com.sun.star.drawing.Shape
.. _drawing_shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Shape.html

.. |drawing_line_shape| replace:: com.sun.star.drawing.LineShape
.. _drawing_line_shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineShape.html

.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _FillProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1FillProperties.html
.. _LineProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineProperties.html
.. _LineShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineShape.html
.. _ParagraphProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html
.. _RotationDescriptor: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1RotationDescriptor.html
.. _ShadowProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1ShadowProperties.html
.. _Shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Shape.html
.. _TextProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1TextProperties.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
.. _LineDash: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1drawing_1_1LineDash.html
.. _LineStyle: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html#a86e0f5648542856159bb40775c854aa7
.. _EllipseShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1EllipseShape.html

.. _ooouno: https://pypi.org/project/ooouno/