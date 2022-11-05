.. _ch13:

********************************
Chapter 13. Drawing Basic Shapes
********************************

.. topic:: Overview

    A Black Dashed Line; A Red Ellipse; Filled Rectangles; Text; Shape Names; A Transparent Circle and a Polar Line; A Math Formula as an OLE Shape; Polygons; Multiple Lines, Partial Ellipses

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

                    s = Draw.draw_formula(
                        slide=curr_slide,
                        formula="func e^{i %pi} + 1 = 0",
                        x=70,
                        y=20,
                        width=75,
                        height=40
                    )
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

|draw_picture_py|_ creates a new Draw document, and finishes by displaying a :ref:`class_msg_box` shown in :numref:`ch13fig_msgbox_all_done` asking the user if they want to close the document.

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
        def make_shape(
            shape_type: DrawingShapeKind | str,
            x: int,
            y: int,
            width: int,
            height: int
        ) -> XShape:

            # parameters are in mm units
            shape = Lo.create_instance_msf(XShape, f"com.sun.star.drawing.{shape_type}", raise_err=True)
            shape.setPosition(Point(x * 100, y * 100))
            shape.setSize(Size(width * 100, height * 100))
            return shape

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
It also check if the (x, y) coordinate is located on the page. If it isn't, :py:meth:`.Draw.warns_position` prints a warning message.

.. tabs::

    .. code-tab:: python

        # in the Draw class (simplified)
        @classmethod
        def add_shape(
            cls,
            slide: XDrawPage,
            shape_type: DrawingShapeKind | str,
            x: int,
            y: int,
            width: int,
            height: int
        ) -> XShape:

            cls.warns_position(slide=slide, x=x, y=y)
            shape = cls.make_shape(shape_type=shape_type, x=x, y=y, width=width, height=height)
            slide.add(shape)
            return shape

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

    There is a `List of all members <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineShape-members.html>`_ link
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

The rectangle example in |draw_gradient_ex|_ comes in seven different colors show in :numref:`ch13fig_seven_fills`.

.. tabs::

    .. code-tab:: python

        # in DrawPicture._draw_Shapes()
        # rectangle with different fills; uses (x, y) width, height
        rect1 = Draw.draw_rectangle(slide=curr_slide, x=70, y=100, width=75, height=25)
        Props.set(rect1, FillColor=CommonColor.LIME)

.. tabs::

    .. code-tab:: python

        # in DrawGradient Class()
        def _gradient_fill(self, curr_slide: XDrawPage) -> None:

            # rectangle shape is also com.sun.star.drawing.FillProperties service
            # casting is only at design time and is not really necessary;
            # however it gives easy access with typing support for other properties
            rect1 = cast(
                "FillProperties",
                Draw.draw_rectangle(
                    slide=curr_slide, x=self._x, y=self._y, width=self._width, height=self._height
                ),
            )
            Props.set(rect1, FillColor=self._start_color)
            # other properties can be set
            # rect1.FillTransparence = 55

.. seealso::

    - :py:meth:`.Draw.draw_rectangle`
    - :py:meth:`.Props.set`

In both |draw_picture|_ and |draw_gradient_ex|_ the code for creating a Rectangle is basically the same.

|draw_gradient_ex|_ demonstrates that ``rect1`` is also a FillProperties_ service
and other properties can be set.

..
    figure 8

.. cssclass:: diagram

    .. _ch13fig_seven_fills:
    .. figure:: https://user-images.githubusercontent.com/4193389/199873235-517287a4-7514-4108-a6a3-2bb6d768e3ca.png
        :alt: Seven Ways of Filling a Rectangle.
        :figclass: align-center

        :Seven Ways of Filling a Rectangle.

13.4.1 Gradient Color
---------------------

``gradient color`` and ``gradient color Custom props`` are actually the same except ``gradient color Custom props``
set properties after the gradient is created.

.. tabs::

    .. code-tab:: python

        # in DrawGradient Class()
        # creates color gradient and color Custom props gradient
        def _gradient_name(self, curr_slide: XDrawPage, set_props: bool) -> None:

            # rectangle shape is also com.sun.star.drawing.FillProperties service
            # casting is only at design time and is not really necessary;
            # however it gives easy access with typing support for other properties
            rect1 = cast(
                "FillProperties",
                Draw.draw_rectangle(
                    slide=curr_slide, x=self._x, y=self._y, width=self._width, height=self._height
                ),
            )
            grad = Draw.set_gradient_color(shape=rect1, name=self._name_gradient)
            if set_props:
                # grad = cast("Gradient", Props.get(rect1, "FillGradient"))
                # print(grad)
                grad.Angle = self._angle * 10  # in 1/10 degree units
                grad.StartColor = self._start_color
                grad.EndColor = self._end_color
                Draw.set_gradient_properties(shape=rect1, grad=grad)
            # rect1.FillTransparence = 40

.. seealso::

    - :py:meth:`.Draw.draw_rectangle`
    - :py:meth:`.Draw.set_gradient_color`

The hardest part of using this function is determining what name value to pass to the ``FillGradientName`` property for FillProperties_ (:abbreviation:`e.g.` "Neon Light").
For this reason |odev| has a :py:class:`~.kind.drawing_gradient_kind.DrawingGradientKind` Enum class that can be passed to :py:meth:`.Draw.set_gradient_color`
for easy lookup of gradient name. Optionally :py:meth:`.Draw.set_gradient_color` can be passed a string name instead of :py:class:`~.kind.drawing_gradient_kind.DrawingGradientKind`.

To see the gradient name fire up Office's Draw application, and check out the gradient names listed in the toolbar.
:numref:`ch13fig_lo_gradient_names` shows what happens when the user selects a shape and chooses the "Gradient" menu item from the combo box.

..
    figure 9

.. cssclass:: screen_shot

    .. _ch13fig_lo_gradient_names:
    .. figure:: https://user-images.githubusercontent.com/4193389/200009116-b3190dbc-4791-4d59-9017-2840edcb87b6.png
        :alt: The Gradient Names in Libre Office.
        :figclass: align-center

        :The Gradient Names in LibreOffice.

Calling ``_gradient_name()`` with ``set_props=True`` will result in creating a gradient similar to ``gradient color Custom props`` of :numref:`ch13fig_seven_fills`.
The actual gradient created will depend on the Properties set for ``DrawGradient`` class instance.

13.4.2 Gradient Common Color
----------------------------

The fourth example in :numref:`ch13fig_seven_fills` shows what happens when you define your own gradient and angle of change. In ``DrawGradient`` class, the call is:

.. tabs::

    .. code-tab:: python

        # in DrawGradient Class()
        # creates gradient CommonColor
        def _gradient(self, curr_slide: XDrawPage) -> None:

            # rectangle shape is also com.sun.star.drawing.FillProperties service
            # casting is only at design time and is not really necessary;
            # however it gives easy access with typing support for other properties
            rect1 = cast(
                "FillProperties",
                Draw.draw_rectangle(
                    slide=curr_slide, x=self._x, y=self._y, width=self._width, height=self._height
                )
            )
            Draw.set_gradient_color(
                shape=rect1,
                start_color=self._start_color,
                end_color=self._end_color,
                angle=Angle(self._angle)
            )
            # rect1.FillTransparence = 40

.. seealso::

    - :py:meth:`.Draw.draw_rectangle`
    - :py:meth:`.Draw.set_gradient_color`

:py:meth:`.Draw.set_gradient_color` has several overloads and calls ``_set_gradient_color_colors()`` internally when setting ``x``, ``y``, ``width`` and ``height`` parameters:

.. tabs::

    .. code-tab:: python

        # from the Draw class (simplified)
        # called by set_gradient_color() overload method
        @classmethod
        def _set_gradient_color_colors(
            cls, shape: XShape, start_color: Color, end_color: Color, angle: Angle
        ) -> Gradient:

            grad = Gradient()
            grad.Style = GradientStyle.LINEAR
            grad.StartColor = start_color
            grad.EndColor = end_color

            grad.Angle = angle.Value * 10  # in 1/10 degree units
            grad.Border = 0
            grad.XOffset = 0
            grad.YOffset = 0
            grad.StartIntensity = 100
            grad.EndIntensity = 100
            grad.StepCount = 10

            cls.set_gradient_properties(shape, grad)

            return Props.get(shape, "FillGradient")


.. seealso::

    - :py:meth:`.Draw.set_gradient_properties`
    - :py:meth:`.Props.get`

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`set_gradient_color`

:py:meth:`.Draw.set_gradient_properties` sets the properties ``FillStyle`` and ``FillGradient``.
The latter requires a Gradient object, which is documented in the FillProperties_ class, as shown in :numref:`ch13fig_api_fill_gradient_prop`.

..
    figure 10

.. cssclass:: screen_shot invert

    .. _ch13fig_api_fill_gradient_prop:
    .. figure:: https://user-images.githubusercontent.com/4193389/200025206-2c169856-3964-4976-bb8c-2db9c998676d.png
        :alt: The Fill Gradient Property in the Fill Properties Class
        :figclass: align-center

        :The ``FillGradient`` Property in the FillProperties_ Class.

Clicking on the ``com::sun:star:awt::Gradient`` name in Figure 10 loads its Gradient_ Struct Reference documentation,
which lists ten fields that need to be set.

The colors passed to :py:meth:`.Draw.set_gradient_color` are :py:data:`.Color` type which is a alias of ``int``.
It is perfectly fine to pass integer values as :py:meth:`.Draw.set_gradient_color` ``start_color`` and ``end_color``

:py:data:`.Color` constants can be found in :py:class:`.color.CommonColor` class.

Example of setting color.

.. tabs::

    .. code-tab:: python

        from ooodev.office.draw import Draw
        from ooodev.utils.color import CommonColor

        # other code ...
        Draw.set_gradient_color(shape=shape, start_color=CommonColor.RED, end_color=CommonColor.GREEN)

13.4.3 Hatching
---------------

The fifth fill in :numref:`ch13fig_seven_fills` employs hatching. In ``DrawGradient`` class, the call is:

.. tabs::

    .. code-tab:: python

        # in DrawGradient Class()
        def _gradient_hatching(self, curr_slide: XDrawPage) -> None:
            # rectangle shape is also com.sun.star.drawing.FillProperties service
            # casting is only at design time and is not really necessary;
            # however it gives easy access with typing support for other properties
            rect1 = cast(
                "FillProperties",
                Draw.draw_rectangle(
                    slide=curr_slide,
                    x=self._x,
                    y=self._y,
                    width=self._width,
                    height=self._height
                ),
            )
            Draw.set_hatch_color(shape=rect1, name=self._hatch_gradient)
            # rect1.FillTransparence = 40


``_gradient_hatching()`` Calls :py:meth:`.Draw.set_hatch_color`.

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @staticmethod
        def set_hatch_color(shape: XShape, name: DrawingHatchingKind | str) -> None:

            props = Lo.qi(XPropertySet, shape, True)
            props.setPropertyValue("FillStyle", FillStyle.HATCH)
            props.setPropertyValue("FillHatchName", str(name))
  

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`set_hatch_color`

This function is much the same as :py:meth:`.Draw.set_gradient_properties` except that it utilizes ``FillHatchName`` rather
than ``FillGradientName``, and the fill style is set to ``FillStyle.HATCH``.
Suitable hatching names can be found by looking at the relevant list in Draw.
:numref:`ch13fig_lo_hatching_names` shows the ``Hatching`` items.

The hardest part of using this function is determining what name value to pass to the ``FillHatchName`` property for FillProperties_ (:abbreviation:`e.g.` "Green 30 Degrees").
For this reason |odev| has a :py:class:`~.kind.drawing_hatching_kind.DrawingHatchingKind` Enum class that can be passed to :py:meth:`.Draw.set_hatch_color`
for easy lookup of gradient name. Optionally :py:meth:`.Draw.set_hatch_color` can be passed a string name instead of :py:class:`~.kind.drawing_hatching_kind.DrawingHatchingKind`.

To see the Hatching names fire up Office's Draw application, and check out the Hatching names listed in the toolbar.
:numref:`ch13fig_lo_hatching_names` shows what happens when the user selects a shape and chooses the "Hatching" menu item from the combo box.

..
    figure 11

.. cssclass:: screen_shot

    .. _ch13fig_lo_hatching_names:
    .. figure:: https://user-images.githubusercontent.com/4193389/200056558-a1d87a3d-db8c-4bf4-8ffe-01718466d030.png
        :alt: The Hatching Names in Libre Office.
        :figclass: align-center

        :The Hatching Names in LibreOffice.

13.4.4 Bitmap Color
-------------------

The sixth rectangle fill in :numref:`ch13fig_seven_fills` utilizes a bitmap color:

.. tabs::

    .. code-tab:: python

        # in DrawGradient Class()
        def _gradient_bitmap(self, curr_slide: XDrawPage) -> None:
            # rectangle shape is also com.sun.star.drawing.FillProperties service
            # casting is only at design time and is not really necessary;
            # however it gives easy access with typing support for other properties
            rect1 = cast(
                "FillProperties",
                Draw.draw_rectangle(
                    slide=curr_slide,
                    x=self._x,
                    y=self._y,
                    width=self._width,
                    height=self._height
                ),
            )
            Draw.set_bitmap_color(shape=rect1, name=self._bitmap_gradient)
            # rect1.FillTransparence = 40

``_gradient_bitmap()`` Calls :py:meth:`.Draw.set_bitmap_color`.

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @staticmethod
        def set_bitmap_color(shape: XShape, name: DrawingBitmapKind | str) -> None:

            props = Lo.qi(XPropertySet, shape, True)
            props.setPropertyValue("FillStyle", FillStyle.BITMAP)
            props.setPropertyValue("FillBitmapName", str(name))

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`set_bitmap_color`

This function is also similar to :py:meth:`.Draw.set_gradient_properties` except that it utilizes ``FillBitmapName`` rather
than ``FillGradientName``, and the fill style is set to ``FillStyle.BITMAP``.
Suitable Bitmap names can be found by looking at the relevant list in Draw.
:numref:`ch13fig_lo_bitmap_names` shows the ``Bitmap`` items.

The hardest part of using this function is determining what name value to pass to the ``FillBitmapName`` property for FillProperties_ (:abbreviation:`e.g.` "Maple Leaves").
For this reason |odev| has a :py:class:`~.kind.drawing_bitmap_kind.DrawingBitmapKind` Enum class that can be passed to :py:meth:`.Draw.set_bitmap_color`
for easy lookup of gradient name. Optionally :py:meth:`.Draw.set_bitmap_color` can be passed a string name instead of :py:class:`~.kind.drawing_bitmap_kind.DrawingBitmapKind`.

To see the Bitmap names fire up Office's Draw application, and check out the Bitmap names listed in the toolbar.
:numref:`ch13fig_lo_bitmap_names` shows what happens when the user selects a shape and chooses the "Bitmap" menu item from the combo box.

..
    figure 12

.. cssclass:: screen_shot

    .. _ch13fig_lo_bitmap_names:
    .. figure:: https://user-images.githubusercontent.com/4193389/200060222-f14cfb7a-8f73-424a-aa4a-ba93fb4ca9b9.png
        :alt: The Bitmap Names in Libre Office
        :figclass: align-center

        :The Bitmap Names in LibreOffice.

13.4.5 Bitmap File Color
------------------------

The final fill in :numref:`ch13fig_seven_fills` loads a bitmap from ``crazy_blue.jpg``:

.. tabs::

    .. code-tab:: python

        # in DrawGradient Class()
        # in this case self._gradient_fnm is crazy_blue.jpg
         def _gradient_bitmap_file(self, curr_slide: XDrawPage) -> None:
            rect1 = Draw.draw_rectangle(
                slide=curr_slide,
                x=self._x,
                y=self._y,
                width=self._width,
                height=self._height
            )
            Draw.set_bitmap_file_color(shape=rect1, fnm=self._gradient_fnm)

``_gradient_bitmap_file()`` Calls :py:meth:`.Draw.set_bitmap_file_color`.

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @staticmethod
        def set_bitmap_file_color(shape: XShape, fnm: PathOrStr) -> None:

            props = Lo.qi(XPropertySet, shape, True)
            props.setPropertyValue("FillStyle", FillStyle.BITMAP)
            props.setPropertyValue("FillBitmapURL", FileIO.fnm_to_url(fnm))

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`set_bitmap_file_color`

The ``FillBitmapURL`` property requires a URL, so the filename is converted by :py:meth:`.FileIO.fnm_to_url`.

13.5 Text
=========

The "Hello LibreOffice" text shape in :numref:`ch13fig_draw_shapes_six` is created by calling :py:meth:`.Draw.draw_text`:

.. tabs::

    .. code-tab:: python

        text1 = Draw.draw_text(
            slide=curr_slide, msg="Hello LibreOffice", x=120, y=120, width=60, height=30, font_size=24
        )
        Props.set(text1, Name="text1")

The first four numerical parameters define the shape's bounding rectangle in terms of its top-left coordinate, width, and height.
The fifth, optional number specifies a font size (in this case, ``24pt``).

:py:meth:`.Draw.draw_text` calls :py:meth:`.Draw.add_shape` with :py:attr:`.DrawingShapeKind.TEXT_SHAPE`:

.. tabs::

    .. code-tab:: python

        # in the draw class (simplified)
        @classmethod
        def draw_text(
            cls,
            slide: XDrawPage,
            msg: str,
            x: int,
            y: int,
            width: int,
            height: int,
            font_size: int = 0
        ) -> XShape:

            shape = cls.add_shape(
                slide=slide,
                shape_type=DrawingShapeKind.TEXT_SHAPE,
                x=x,
                y=y,
                width=width,
                height=height
            )
            cls.add_text(shape=shape, msg=msg, font_size=font_size)
            return shape

:py:meth:`~.Draw.add_shape` adds the message to the shape, and sets its font size:

.. tabs::

    .. code-tab:: python

        # in the draw class (simplified)
        @classmethod
        def add_shape(
            cls,
            slide: XDrawPage,
            shape_type: DrawingShapeKind | str,
            x: int,
            y: int,
            width: int,
            height: int
        ) -> XShape:

            cls.warns_position(slide=slide, x=x, y=y)
            shape = cls.make_shape(shape_type=shape_type, x=x, y=y, width=width, height=height)
            slide.add(shape)
            return shape

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`draw_text`
        -  :odev_src_draw_meth:`add_shape`

The shape is converted into an XText_ reference, and the text range selected with a cursor.

The ``CharHeight`` property comes from the CharacterProperties_ service, which is inherited by the Text_ service (as shown in  :numref:`ch11fig_rectangel_shape_props`).

Some Help with Text Properties
------------------------------

The text-related properties for a shape can be accessed with :py:meth:`.Draw.get_text_properties`:

.. tabs::

    .. code-tab:: python

        # in the draw class (simplified)
        @staticmethod
        def get_text_properties(shape: XShape) -> XPropertySet:
            xtxt = Lo.qi(XText, shape, True)
            cursor = xtxt.createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            xrng = Lo.qi(XTextRange, cursor, True)
            return Lo.qi(XPropertySet, xrng, True)


.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`get_text_properties`

``_draw_shapes()`` in |draw_picture_py|_ calls :py:meth:`.Draw.get_text_properties` on the ``text1`` TextShape_, and prints all its properties:

.. tabs::

    .. code-tab:: python

        # in _draw_shapes() in draw_picture.py
        Props.show_props("TextShape's Text Properties", Draw.get_text_properties(text1))

The output is long, but includes the line:

::

  CharHeight = 24.0

which indicates that the font size was correctly changed by the earlier call to :py:meth:`.Draw.draw_text`.

13.6 Using a Shape Name
=======================

Immediately after the call to :py:meth:`.Draw.draw_text`, the shape's name is set:

.. tabs::

    .. code-tab:: python

        # in _draw_shapes() in draw_picture.py
        Props.set(text1, Name="text1")

The ``Name`` property, which is defined in the Shape_ class, is a useful way of referring to a shape.
The ``show()`` function of |draw_picture_py|_ passes a name to :py:meth:`.Draw.find_shape_by_name`:

.. tabs::

    .. code-tab:: python

        # in show() in draw_picture.py
        s = Draw.find_shape_by_name(curr_slide, "text1")
        Draw.report_pos_size(s)

.. tabs::

    .. code-tab:: python

        # in the draw class (simplified)
        @classmethod
        def find_shape_by_name(cls, slide: XDrawPage, shape_name: str) -> XShape:
            shapes = cls.get_shapes(slide)
                sn = shape_name.casefold()
            if not shapes:
                raise ShapeMissingError("No shapes were found in the draw page")

            for shape in shapes:
                nm = str(Props.get(shape, "Name")).casefold()
                if nm == sn:
                    return shape
            raise mEx.ShapeMissingError(f'No shape named "{shape_name}"')

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`find_shape_by_name`

:py:meth:`.Draw.get_shapes` builds a list of shapes by iterating through the XDrawPage object as an indexed container of shapes:

In this case :py:meth:`.Draw.get_shapes` call the internal Draw method ``_get_shapes_slide()``.

.. tabs::

    .. code-tab:: python

        # in the draw class (simplified)
        @classmethod
        def _get_shapes_slide(cls, slide: XDrawPage) -> List[XShape]:
            if slide.getCount() == 0:
                return []

            shapes: List[XShape] = []
            for i in range(slide.getCount()):
                shapes.append(mLo.Lo.qi(XShape, slide.getByIndex(i), True))
            return shapes

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`get_shapes`

:py:meth:`.Draw.report_pos_size` prints some brief information about a shape, including its name, shape type, position, and size:

.. tabs::

    .. code-tab:: python

        # in the draw class
        @classmethod
        def report_pos_size(cls, shape: XShape) -> None:
            if shape is None:
                print("The shape is null")
                return
            print(f'Shape Name: {Props.get(shape, "Name")}')
            print(f"  Type: {shape.getShapeType()}")
            cls.print_point(shape.getPosition())
            cls.print_size(shape.getSize())

``XShape.getShapeType()`` returns the class name of the shape as a string (in this case, TextShape_).

13.7 A Transparent Circle and a Polar Line
==========================================

The last two shapes created by |draw_picture_py|_ ``_draw_shapes()`` are a gray transparent circle and a polar line.

.. tabs::

    .. code-tab:: python

        # in _draw_shapes() in draw_picture.py
        # gray transparent circle; uses (x,y), radius
        circle2 = Draw.draw_circle(slide=curr_slide, x=40, y=150, radius=20)
        Props.set(circle2, FillColor=CommonColor.GRAY)
        Draw.set_transparency(shape=circle2, level=Intensity(25))

        # thick line; uses (x,y), angle clockwise from x-axis, length
        line2 = Draw.draw_polar_line(slide=curr_slide, x=60, y=200, degrees=45, distance=100)
        Props.set(line2, LineWidth=300)

A polar line is one defined using polar coordinates, which specifies the coordinate of one end of the line,
and the angle and length of the line from that point.

:py:meth:`.Draw.draw_circle` uses an EllipseShape_, and :py:meth:`.Draw.draw_polar_line` converts the polar values into two coordinates so :py:meth:`.Draw.draw_line` can be called.

13.8 A Math formula as an OLE Shape
===================================

.. todo::

    Chapter 13.8, Add link to part 5

Draw/Impress documents can include OLE objects through ``OLE2Shape``, which allows a shape to link to an external document.
Probably the most popular kind of OLE shape is the chart, we will have a detailed discussion of that topic when we get to Part 5, although there is a code snippet below.

The best way of finding out what OLE objects are available is to go to Draw's (or Impress') Insert menu, Object, "OLE Object" dialog.
It lists Office spreadsheet, chart, drawing, presentation, and formula documents, and a range of Microsoft and PDF types (when you click on "Further objects").

The |draw_picture|_ OLE example displays a mathematical formula, as in :numref:`ch12fig_draw_math_formula`.

..
    Figure 13

.. cssclass:: diagram invert

    .. _ch12fig_draw_math_formula:
    .. figure:: https://user-images.githubusercontent.com/4193389/200079304-62bbd25c-4e69-4cdb-9dac-65e58bbedc3d.png
        :alt: A Math Formula in a Draw Document.
        :figclass: align-center

        :A Math Formula in a Draw Document.

|draw_picture_py|_ renders the formula by calling :py:meth:`.Draw.draw_formula`, which hides the tricky aspects of instantiating the OLE shape:

.. tabs::

    .. code-tab:: python

        # in show() in draw_picture.py
        s = Draw.draw_formula(
            slide=curr_slide,
            formula="func e^{i %pi} + 1 = 0",
            x=70,
            y=20,
            width=75,
            height=40
        )

The second argument is a formula string, written using Office's Math notation.
For an overview, see the "Commands Reference" appendix of the "Math Guide", available from https://libreoffice.org/get-help/documentation.

:py:meth:`.Draw.draw_formula` is coded as:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def draw_formula(
            cls,
            slide: XDrawPage,
            formula: str,
            x: int,
            y: int,
            width: int,
            height: int
        ) -> XShape:

            shape = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.OLE2_SHAPE, x=x, y=y, width=width, height=height
            )
            cls.set_shape_props(shape, CLSID=str(Lo.CLSID.MATH))  # a formula

            model = mLo.Lo.qi(XModel, Props.get(shape, "Model"), True)
            # Info.show_services(obj_name="OLE2Shape Model", obj=model)
            Props.set(model, Formula=formula)

            # for some reason setting model Formula here cause the shape size to be blown out.
            # resetting size and positon corrects the issue.
            cls.set_size(shape, Size(width, height))
            cls.set_position(shape, Point(x, y))
            return shape

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`draw_formula`

``OLE2Shape`` uses a ``CLSID`` property to hold the class ID of the OLE object.
Setting this property affects the shape's model (data format), which is stored in the ``Model`` property.
:py:meth:`~.Draw.draw_formula` casts this property to XModel_ and, since the model represents formula data,
it has a ``Formula`` property where the formula string is stored.

Creating Other Kinds of OLE Shape
---------------------------------

The use of a ``Formula`` property in :py:meth:`.Draw.draw_formula` only works for an OLE shape representing a formula. How are other kinds of data stored?

The first step is to set the OLE shape's class ID to the correct value, which will affect its ``Model`` property.
:py:class:`.Lo.CLSID` is an enum containing some of the class ID's.
Note its use in the previous code example, ``cls.set_shape_props(shape, CLSID=str(Lo.CLSID.MATH))``.

Creating an OLE2Shape for a chart begins like so:

.. tabs::

    .. code-tab:: python

        shape = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.OLE2_SHAPE, x=x, y=y, width=width, height=height
            )
        cls.set_shape_props(shape, CLSID=str(Lo.CLSID.CHART_CLSID))
        model = Lo.qi(XModel, Props.get(shape, "Model"))

Online information on how to use XModel_ to store a chart, a graphic, or something else, is pretty sparse.
A good way is to list the services that support the XModel_ reference. This is done by calling :py:meth:`.Info.show_services`:

.. tabs::

    .. code-tab:: python

        Info.show_services("OLE2Shape Model", model)

For the version of model in :py:meth:`~.Draw.draw_formula`, it reports:

::

    OLE2Shape Model Supported Services (2)
      "com.sun.star.document.OfficeDocument"
      "com.sun.star.formula.FormulaProperties"

This gives a strong hint to look inside the FormulaProperties_ service, to find a property for storing the formula string.
A look at the documentation reveals a ``Formula`` property, which is used in :py:meth:`~.Draw.draw_formula`.

When the model refers to chart data, the same call to :py:meth:`.Info.show_services` prints:

::

    OLE2Shape Model Supported Services (3)
      "com.sun.star.chart.ChartDocument"
      "com.sun.star.chart2.ChartDocument"
      "com.sun.star.document.OfficeDocument"

.. todo::

    Chapter 13.8, Add link to Part 5.

The ``com.sun.star.chart2`` package is the newer way of manipulating charts, which suggests that the XModel_ interfaces should be converted to an interface of ``com.sun.star.chart2.ChartDocument``.
The most useful is XChartDocument_, which is obtained via: ``chart_doc = Lo.qi(XChartDocument, model)`` XChartDocument_ supports a rich set of chart manipulation methods.
We'll return to charts in Part 5.

13.9 Polygons
=============

The main() function of |animate_bike_py|_ calls :py:meth:`.Draw.draw_polygon` twice to create regular polygons for a square and pentagon:


.. tabs::

    .. code-tab:: python

        # in animate() of anim_bicycle.py
        square = Draw.draw_polygon(slide=slide, x=125, y=125, sides=PolySides(4), radius=25)
        Props.set(square, FillColor=CommonColor.LIGHT_GREEN)

        pentagon = Draw.draw_polygon(slide=slide, x=150, y=75, sides=PolySides(5))
        Props.set(pentagon, FillColor=CommonColor.PURPLE)

The polygons can be seen in :numref:`ch12fig_bike_and_shapes`.

..
    Figure 14

.. cssclass:: screen_shot invert

    .. _ch12fig_bike_and_shapes:
    .. figure:: https://user-images.githubusercontent.com/4193389/200084815-fcb643b3-7044-40b6-8cd4-26094799418c.png
        :alt: Bicycle and Shapes
        :figclass: align-center

        :Bicycle and Shapes.

:py:meth:`.Draw.draw_polygon` is:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def draw_polygon(
            cls,
            slide: XDrawPage,
            x: int,
            y: int,
            sides: PolySides,
            radius: int = POLY_RADIUS
        ) -> XShape:

            if radius is None:
                radius = Draw.POLY_RADIUS
            polygon = cls.add_shape(
                slide=slide,
                shape_type=DrawingShapeKind.POLY_POLYGON_SHAPE,
                x=0,
                y=0,
                width=0,
                height=0,
            )
            pts = cls.gen_polygon_points(x=x, y=y, radius=radius, sides=sides)
            # could be many polygons pts in this 2D array
            polys = (pts,)
            prop_set = mLo.Lo.qi(XPropertySet, polygon, raise_err=True)
            polyseq = uno.Any("[][]com.sun.star.awt.Point", polys)
            uno.invoke(prop_set, "setPropertyValue", ("PolyPolygon", polyseq))
            return polygon

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`draw_polygon`

:py:meth:`~.Draw.draw_polygon` creates a PolyPolygonShape_ shape which is designed to store multiple polygons.
This is why the polys data structure instantiated at the end of :py:meth:`~.Draw.draw_polygon` is an array of points arrays,
since the shape's ``PolyPolygon`` property can hold multiple point arrays. However, :py:meth:`~.Draw.draw_polygon` only creates
a single points array by calling :py:meth:`~.Draw.gen_polygon_points`.

A points array defining the four points of a square could be:

.. tabs::

    .. code-tab:: python

        from ooo.dyn.awt.point import Point

        pts (
            Point(4_000, 1_200),
            Point(4_000, 2_000),
            Point(5_000, 2_000),
            Point(5_000, 1_200)
        )

.. note::

    The coordinates of each point use Office's ``1/100 mm`` units.

:py:meth:`~.Draw.gen_polygon_points` generates a points array for a regular polygon based on the coordinate of the center of the polygon,
the distance from the center to each point (the shape's radius), and the required number of sides:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @staticmethod
        def gen_polygon_points(x: int, y: int, radius: int, sides: PolySides) -> Tuple[Point, ...]:

            pts: List[Point] = []
            angle_step = math.pi / sides.Value
            for i in range(sides.Value):
                pt = Point(
                    int(round(((x * 100) + ((radius * 100)) * math.cos(i * 2 * angle_step)))),
                    int(round(((y * 100) + ((radius * 100)) * math.sin(i * 2 * angle_step)))),
                )
                pts.append(pt)
            return tuple(pts)

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`gen_polygon_points`

13.10 Multi-line Shapes
=======================

A PolyLineShape_ can hold multiple line paths, where a path is a sequence of connected lines.
:py:meth:`.Draw.draw_lines` only creates a single line path, based on being passed arrays of ``x-`` and ``y-`` axis coordinates.
For example, the following code in |animate_bike_py|_ creates the crossed lines at the top-left of :numref:`ch12fig_bike_and_shapes`:

.. tabs::

    .. code-tab:: python

        # in animate() of anim_bicycle.py

        xs = (10, 30, 10, 30)
        ys = (10, 100, 100, 10)
        Draw.draw_lines(slide=slide, xs=xs, ys=ys)

:py:meth:`.Draw.draw_lines` is:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def draw_lines(cls, slide: XDrawPage, xs: Sequence[int], ys: Sequence[int]) -> XShape:

            num_points = len(xs)
            if num_points != len(ys):
                raise IndexError("xs and ys must be the same length")

            pts: List[Point] = []
            for x, y in zip(xs, ys):
                # in 1/100 mm units
                pts.append(Point(x * 100, y * 100))

            # an array of Point arrays, one Point array for each line path
            line_paths = (tuple(pts),)

            # for a shape formed by from multiple connected lines
            poly_line = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.POLY_LINE_SHAPE, x=0, y=0, width=0, height=0
            )
            prop_set = mLo.Lo.qi(XPropertySet, poly_line, raise_err=True)
            seq = uno.Any("[][]com.sun.star.awt.Point", line_paths)
            uno.invoke(prop_set, "setPropertyValue", ("PolyPolygon", seq))
            return poly_line

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`draw_lines`

:py:meth:`~.Draw.draw_lines` creates an tuple of Point tuples which is stored in the PolyLineShape_ property called ``PolyPolygon``.
However, :py:meth:`~.Draw.draw_lines` only adds a single points tuple to the ``line_paths`` data structure since only one line path is being created.

13.11 Partial Ellipses
======================

EllipseShape_ contains a ``CircleKind`` property that determines whether the entire ellipse should be drawn, or only a portion of it.
The properties ``CircleStartAngle`` and ``CircleEndAngle`` define the angles where the solid part of the ellipse starts and finishes.
Zero degrees is the positive ``x-axis``, and the angle increase in ``1/100`` degrees units counter-clockwise around the center of the ellipse.


|animate_bike_py|_ contains the following example:

.. tabs::

    .. code-tab:: python

        # in animate() of anim_bicycle.py
        pie = Draw.draw_ellipse(slide=slide, x=30, y=slide_size.Width - 100, width=40, height=20)
        Props.set(
            pie,
            FillColor=CommonColor.LIGHT_SKY_BLUE,
            CircleStartAngle=9_000,  #   90 degrees ccw
            CircleEndAngle=36_000,  #    360 degrees ccw
            CircleKind=CircleKind.SECTION,
        )

This creates the blue partial ellipse shown at the bottom left of :numref:`ch12fig_bike_and_shapes`.

:numref:`ch12fig_partial_ellipses` shows the different results when CircleKind_ is set to ``CircleKind.SECTION``, ``CircleKind.CUT``, and ``CircleKind.ARC``.

..
    Figure 15

.. cssclass:: screen_shot invert

    .. _ch12fig_partial_ellipses:
    .. figure:: https://user-images.githubusercontent.com/4193389/200087984-72de9e74-6654-4263-a6fa-088db523207a.png
        :alt: Different Types of Partial Ellipse
        :figclass: align-center

        :Different Types of Partial Ellipse

.. |animate_bike| replace:: Animate Bike
.. _animate_bike: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_animate_bike

.. |animate_bike_py| replace:: anim_bicycle.py
.. _animate_bike_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/draw/odev_animate_bike/anim_bicycle.py

.. |draw_picture| replace:: Draw Picture
.. _draw_picture: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_draw_picture

.. |draw_picture_py| replace:: draw_picture.py
.. _draw_picture_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_draw_picture/draw_picture.py

.. |drawing_shape| replace:: com.sun.star.drawing.Shape
.. _drawing_shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Shape.html

.. |drawing_line_shape| replace:: com.sun.star.drawing.LineShape
.. _drawing_line_shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineShape.html

.. |draw_gradient_ex| replace:: Draw Gradient Examples
.. _draw_gradient_ex: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_gradient

.. _CharacterProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
.. _CircleKind: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html#a6a52201f72a50075b45fea2c19340c0e
.. _EllipseShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1EllipseShape.html
.. _FillProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1FillProperties.html
.. _FormulaProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1formula_1_1FormulaProperties.html
.. _Gradient: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1awt_1_1Gradient.html
.. _LineDash: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1drawing_1_1LineDash.html
.. _LineProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineProperties.html
.. _LineShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineShape.html
.. _LineStyle: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html#a86e0f5648542856159bb40775c854aa7
.. _ParagraphProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html
.. _PolyLineShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1PolyLineShape.html
.. _PolyPolygonShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1PolyPolygonShape.html
.. _RotationDescriptor: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1RotationDescriptor.html
.. _ShadowProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1ShadowProperties.html
.. _Shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Shape.html
.. _Text: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1Text.html
.. _TextProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1TextProperties.html
.. _TextShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1TextShape.html
.. _XChartDocument: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1chart2_1_1XChartDocument.html
.. _XModel: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XModel.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html

.. _ooouno: https://pypi.org/project/ooouno/