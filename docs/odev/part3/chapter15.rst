.. _ch15:

**************************
Chapter 15. Complex Shapes
**************************

.. topic:: Overview

    Connecting Two Rectangles; Shape Composition (grouping, binding, and combining); Combining with Dispatches; Undoing Composition; Bezier Curves (simple and complex)

    Examples: |grouper|_, |draw_bezier_curve|_

This chapter looks at three complex topics involving shapes: connecting rectangles, shape composition, and drawing Bezier curves.

.. _ch15_connecting_tow_rectangles:

15.1 Connecting Two Rectangles
==============================

A line can be drawn between two shapes using a LineShape_.
But it's much easier to join the shapes with a ConnectorShape_, which can be attached precisely by linking its two ends to glue points on the shapes.
Glue points are the little blue circles which appear on a shape when you use connectors in Draw.
They occur in the middle of the upper, lower, left, and right sides of the shape, although it's possible to create extra ones.

By default a connector is drawn using multiple horizontal and vertical lines.
It's possible to change this to a curve, a single line, or a connection made up of multiple lines which are mostly horizontal and vertical.
:numref:`ch15fig_styles_of_connector` shows the four types linking the same two rectangles.

..
    figure 1
    Orig: https://user-images.githubusercontent.com/4193389/200185083-a6e76a2c-a5c1-41b4-b587-52fbe9c8f632.png

.. cssclass:: diagram

    .. _ch15fig_styles_of_connector:
    .. figure:: https://user-images.githubusercontent.com/4193389/200192580-3656ef6a-5a21-49fd-b752-fb39f38f9c2a.png
        :alt: Different Styles of Connector.
        :figclass: align-center

        :Different Styles of Connector.

|grouper_py|_ contains code for generating the top-left example in :numref:`ch15fig_styles_of_connector`:

.. tabs::

    .. code-tab:: python

        # partial main() in grouper.py
        def main(self) -> None:
            loader = Lo.load_office(Lo.ConnectPipe())

            try:
                doc = Draw.create_draw_doc(loader)
                GUI.set_visible(is_visible=True, odoc=doc)
                Lo.delay(1_000)  # need delay or zoom may not occur
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                curr_slide = Draw.get_slide(doc=doc, idx=0)

                print()
                print("Connecting rectangles ...")
                g_styles = Info.get_style_container(doc=doc, family_style_name="graphics")
                # Info.show_container_names("Graphic styles", g_styles)

                self._connect_rectangles(slide=curr_slide, g_styles=g_styles)

                # code for grouping, binding, and combining shape,
                # discussed later

                Lo.delay(1_500)
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

The ``_connect_rectangles()`` function creates two labeled rectangles, and links them with a standard connector.
The connector starts on the bottom edge of the green rectangle and finishes at the top edge of the blue one (as shown in the top-left of :numref:`ch15fig_styles_of_connector`).
The method also prints out some information about the glue points of the blue rectangle.

.. tabs::

    .. code-tab:: python

        # _connect_rectangles() from grouper.py
        def _connect_rectangles(self, slide: XDrawPage, g_styles: XNameContainer) -> None:
            # draw two two labelled rectangles, one green, one blue, and
            #  connect them. Changing the connector to an arrow

            # dark green rectangle with shadow and text
            green_rect = Draw.draw_rectangle(slide=slide, x=70, y=180, width=50, height=25)
            Props.set(green_rect, FillColor=CommonColor.DARK_SEA_GREEN, Shadow=True)
            Draw.add_text(shape=green_rect, msg="Green Rect")

            # (blue, the default color) rectangle with shadow and text
            blue_rect = Draw.draw_rectangle(slide=slide, x=140, y=220, width=50, height=25)
            Props.set(blue_rect, Shadow=True)
            Draw.add_text(shape=blue_rect, msg="Blue Rect")

            # connect the two rectangles; from the first shape to the second
            conn_shape = Draw.add_connector(
                slide=slide,
                shape1=green_rect,
                shape2=blue_rect,
                start_conn=GluePointsKind.BOTTOM,
                end_conn=GluePointsKind.TOP,
            )

Note that :py:meth:`.Draw.add_text` is used to label the shapes.

:py:meth:`.Draw.add_connector` links the two rectangles based on glue point names supplied as arguments ``start_conn`` and ``end_conn``.
These names are defined in the :py:class:`~.kind.glue_points_kind.GluePointsKind` enum.

:py:meth:`.Draw.add_connector` creates a ConnectorShape_ object and sets several of its properties.
A simplified inheritance hierarchy for ConnectorShape_ is shown in :numref:`ch15fig_connector_shape_hierarchy`, with the parts important for connectors drawn in red.

..
    figure 2

.. cssclass:: diagram invert

    .. _ch15fig_connector_shape_hierarchy:
    .. figure:: https://user-images.githubusercontent.com/4193389/200186952-bcfdb33f-c3c1-4e7b-b2e8-bedf4230866a.png
        :alt: The Connector Shape Hierarchy
        :figclass: align-center

        :The ConnectorShape_ Hierarchy.

Unlike many shapes, such as the RectangleShape_, ConnectorShape_ doesn't have a FillProperties class;
instead it has ConnectorProperties_ which holds most of the properties used by :py:meth:`.Draw.add_connector` which is defined as:

.. tabs::

    .. code-tab:: python

        # in Draw Class (simplified)
        @classmethod
        def add_connector(
            cls,
            slide: XDrawPage,
            shape1: XShape,
            shape2: XShape,
            start_conn: GluePointsKind = None,
            end_conn: GluePointsKind = None,
        ) -> XShape:
            if start_conn is None:
                start_conn = GluePointsKind.RIGHT
            if end_conn is None:
                end_conn = GluePointsKind.LEFT

            xconnector = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.CONNECTOR_SHAPE, x=0, y=0, width=0, height=0
            )
            prop_set = Lo.qi(XPropertySet, xconnector, True)
            prop_set.setPropertyValue("StartShape", shape1)
            prop_set.setPropertyValue("StartGluePointIndex", int(start_conn))

            prop_set.setPropertyValue("EndShape", shape2)
            prop_set.setPropertyValue("EndGluePointIndex", int(end_conn))

            prop_set.setPropertyValue("EdgeKind", ConnectorType.STANDARD)
            return xconnector

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`add_connector`

:py:meth:`.Draw.add_shape` is called with a (0,0) position, zero width and height.
The real position and dimensions of the connector are set via its properties.
``StartShape`` and ``StartGluePointIndex`` specify the starting shape and its glue point, and ``EndShape`` and ``EndGluePointIndex`` define the ending shape and its glue point.
``EdgeKind`` specifies one of the connection types from :numref:`ch15fig_styles_of_connector`.

|grouper_py|_'s ``_connect_rectangles()`` has some code for retrieving an array of glue points for a shape:

.. tabs::

    .. code-tab:: python

        # _connect_rectangles() from grouper.py
        gps = Draw.get_glue_points(blue_rect)

:py:meth:`.Draw.get_glue_points` converts the shape into an XGluePointsSupplier_, and calls its ``getGluePoints()`` method to retrieves a tuple of GluePoint2_ objects.
To simplify the access to the points data, this structure is returned as a tuple:

.. tabs::

    .. code-tab:: python

        # in Draw Class (simplified)
        @staticmethod
        def get_glue_points(shape: XShape) -> Tuple[GluePoint2, ...]:
            gp_supp = mLo.Lo.qi(XGluePointsSupplier, shape, True)
            glue_pts = gp_supp.getGluePoints()

            num_gps = glue_pts.getCount()  # should be 4 by default
            if num_gps == 0:
                return ()

            gps: List[GluePoint2] = []
            for i in range(num_gps):
                try:
                    gps.append(glue_pts.getByIndex(i))
                except Exception as e:
                    mLo.Lo.print(f"Could not access glue point: {i}")
                    mLo.Lo.print(f"  {e}")

            return tuple(gps)

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`get_glue_points`

``_connect_rectangles()`` doesn't do much with this data, aside from printing out each glue points coordinate.
They're specified in ``1/100 mm`` units relative to the center of the shape.

:numref:`ch15fig_styles_of_connector` shows that connectors don't have arrows, but this can be remedied by changing the connector's graphics style.
The ``graphics`` style family is obtained by :py:meth:`.Info.get_style_container`, and passed to ``_connect_rectangles()``:


.. tabs::

    .. code-tab:: python

        # in main() of grouper.py
        g_styles = Info.get_style_container(doc=doc, family_style_name="graphics")
        self._connect_rectangles(slide=curr_slide, g_styles=g_styles)


The styles reported by :py:meth:`.Info.get_style_container` are related to the Draw built in styles seen in :numref:`ch15fig_ss_line_style`.


.. cssclass:: screen_shot invert

    .. _ch15fig_ss_line_style:
    .. figure:: https://user-images.githubusercontent.com/4193389/200191071-70e283e2-314f-4f70-ba01-50cae278d7dc.png
        :alt: Draw Lines Styles
        :figclass: align-center

        :Draw Lines Styles

Inside ``_connect_rectangles()``, the connector's graphic style is changed to use arrows:

.. tabs::

    .. code-tab:: python

        # in _connect_rectangles() of grouper.py
        Draw.set_style(shape=conn_shape, graphic_styles=g_styles, style_name=GraphicStyleKind.ARROW_LINE)

The :py:attr:`.GraphicStyleKind.ARROW_LINE` style creates black arrows as seen in :numref:`ch15fig_connector_with_arrows`.

..
    figure 3

.. cssclass:: diagram

    .. _ch15fig_connector_with_arrows:
    .. figure:: https://user-images.githubusercontent.com/4193389/200198977-3af76305-cc52-425d-a0f9-542b81c20d0c.png
        :alt: A Connector with an Arrows.
        :figclass: align-center

        :A Connector with an Arrows.

The line width can be adjusted by setting the shape's ``LineWidth`` property (which is defined in the LineProperties_ class), and its color with ``LineColor``.
The result can be seen in :numref:`ch15fig_connector_with_orange_line_arrow`.

.. tabs::

    .. code-tab:: python

        # in _connect_rectangles() of grouper.py
        Props.set(
            conn_shape,
            LineWidth=50,
            LineColor=CommonColor.DARK_ORANGE,
            LineStartName=str(GraphicArrowStyleKind.ARROW_SHORT),
            LineStartCenter=False,
            LineEndName=GraphicArrowStyleKind.NONE,
        )

.. cssclass:: diagram

    .. _ch15fig_connector_with_orange_line_arrow:
    .. figure:: https://user-images.githubusercontent.com/4193389/200199323-66e6e62d-169b-4123-8457-d10ec7011ad0.png
        :alt: An orange line connector with a single arrow.
        :figclass: align-center

        :An orange line connector with a single arrow.

The arrow head can be modified by changing the arrow name assigned to the connector's ``LineStartName`` property, and by setting ``LineStartCenter`` to false.
The place to find names for arrow heads is the Line dialog box in LibreOffice's "Line and Filling" toolbar.
The names appear in the "Start styles" combo-box, as shown in :numref:`ch15fig_arrow_styles`.

..
    figure 4

.. cssclass:: screen_shot invert

    .. _ch15fig_arrow_styles:
    .. figure:: https://user-images.githubusercontent.com/4193389/200199936-ee47b66f-0f12-4dfd-9af0-4074b9df195c.png
        :alt: The Arrow Styles in Libre Office
        :figclass: align-center

        :The Arrow Styles in LibreOffice.

|odev| has :py:class:`~.kind.graphic_arrow_style_kind.GraphicArrowStyleKind` for looking up arrow name to make this task much easier.

If the properties are set to:

.. tabs::

    .. code-tab:: python

        # in _connect_rectangles() of grouper.py
        Props.set(
            conn_shape,
            LineWidth=50,
            LineColor=CommonColor.PURPLE,
            LineStartName=str(GraphicArrowStyleKind.LINE_SHORT),
            LineStartCenter=False,
            LineEndName=GraphicArrowStyleKind.NONE,
        )

then the arrow head changes to that shown in :numref:`ch15fig_arrow_line_purple`.

..
    figure 5

.. cssclass:: diagram

    .. _ch15fig_arrow_line_purple:
    .. figure:: https://user-images.githubusercontent.com/4193389/200200403-20008961-07b3-4459-990e-9afc4dd5f790.png
        :alt: A Different Arrow
        :figclass: align-center

        :A Different Arrow

An arrow can be added to the other end of the connector by adjusting its ``LineEndCenter`` and ``LineEndName`` properties.

|odev| has :py:class:`~.kind.graphic_style_kind.GraphicStyleKind` that makes it much easier to get the ``style_name`` to pass
to :py:meth:`.Draw.set_style`. Styles can be lookec up in the following mannor:

.. tabs::

    .. code-tab:: python

        g_styles = Info.get_style_container(doc=doc, family_style_name="graphics")
        Info.show_container_names("Graphic styles", g_styles)

Alternatively, you can browse through the LineProperties class inherited by ConnectorShape (shown in :numref:`ch15fig_connector_shape_hierarchy`).

15.2 Shape Composition
======================

Office supports three kinds of shape composition for converting multiple shapes into a single shape.
The new shape is automatically added to the page, and the old shapes are removed.
The three techniques are:

1. grouping: the shapes form a single shape without being changed in any way. Office has two mechanisms for grouping: the ``ShapeGroup`` shape and the deprecated XShapeGrouper_ interface;
2. binding: this is similar to grouping, but also draws connector lines between the original shapes;
3. combining: the new shape is built by changing the original shapes if they overlap each other. Office supports four combination styles, called merging, subtraction, intersection, and combination (the default).

|grouper_py|_ illustrates these techniques:

.. tabs::

    .. code-tab:: python

        # partial main() in grouper.py
        # ...
        slide_size = Draw.get_slide_size(curr_slide)
        width = 40
        height = 20
        x = round(((slide_size.Width * 3) / 4) - (width / 2))
        y1 = 20
        y2 = round((slide_size.Height / 2) - (y1 + height))  # so separated
        # y2 = 30  # so overlapping

        s1 = Draw.draw_ellipse(slide=curr_slide, x=x, y=y1, width=width, height=height)
        s2 = Draw.draw_ellipse(slide=curr_slide, x=x, y=y2, width=width, height=height)

        Draw.show_shapes_info(curr_slide)

        # group, bind, or combine the ellipses
        print()
        print("Grouping (or binding) ellipses ...")
        # self._group_ellipses(slide=curr_slide, s1=s1, s2=s2)
        # self._bind_ellipses(slide=curr_slide, s1=s1, s2=s2)
        self._combine_ellipses(slide=curr_slide, s1=s1, s2=s2)
        Draw.show_shapes_info(curr_slide)

        # combine some rectangles
        comp_shape = self._combine_rects(doc=doc, slide=curr_slide)
        Draw.show_shapes_info(curr_slide)

        print("Waiting a bit before splitting...")
        Lo.delay(3000)  # delay so user can see previous composition
        # ...

Two ellipses are created, and positioned at the top-right of the page.

:py:meth:`.Draw.show_shapes_info` is called to supply information about all the shapes on the page:

::

    Draw Page shapes:
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 0
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 1
      Shape service: com.sun.star.drawing.ConnectorShape; z-order: 2
      Shape service: com.sun.star.drawing.EllipseShape; z-order: 3
      Shape service: com.sun.star.drawing.EllipseShape; z-order: 4

The two rectangles and the connector listed first are the results of calling ``_connect_rectangles()`` earlier |grouper_py|_.
The two ellipses were just created in the code snipper given above.

15.2.1 Grouping Shapes
----------------------

|grouper_py|_ calls ``_group_ellipses()`` to group the two ellipses:

.. tabs::

    .. code-tab:: python

        # Grouper.main() of grouper.py
        s1 = Draw.draw_ellipse(slide=curr_slide, x=x, y=y1, width=width, height=height)
        s2 = Draw.draw_ellipse(slide=curr_slide, x=x, y=y2, width=width, height=height)
        self._group_ellipses(slide=curr_slide, s1=s1, s2=s2)

``_group_ellipses()`` is:

.. tabs::

    .. code-tab:: python

        # in Grouper class of grouper.py
        def _group_ellipses(self, slide: XDrawPage, s1: XShape, s2: XShape) -> None:
            shape_group = Draw.add_shape(
                slide=slide, shape_type=DrawingShapeKind.GROUP_SHAPE, x=0, y=0, width=0, height=0
            )
            shapes = Lo.qi(XShapes, shape_group, True)
            shapes.add(s1)
            shapes.add(s2)

The GroupShape_ is converted to an XShapes_ interface so the two ellipses can be added to it.
Note that GroupShape_ has no position or size; they are determined from the added shapes.

An alternative approach for grouping is the deprecated XShapeGrouper_, but it requires a few more lines of coding.
An example can be found in the Developer's Guide, at https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Grouping,_Combining_and_Binding

The on-screen result of ``_group_ellipses()`` is that the two ellipses become a single shape, as poorly shown in :numref:`ch15fig_grouped_ellipses`.

Run the |grouper|_ example with these args.

.. code::

    python -m start -k group

..
    figure 6

.. cssclass:: screen_shot invert

    .. _ch15fig_grouped_ellipses:
    .. figure:: https://user-images.githubusercontent.com/4193389/200201985-4d98ff3a-5db2-4a03-b828-f0ca95ffa211.png
        :alt: The Grouped Ellipses.
        :figclass: align-center

        :The Grouped Ellipses.

There's no noticeable difference from two ellipses until you click on one of them, which causes both to be selected as a single shape.

The change is better shown by a second call to :py:meth:`.Draw.show_shapes_info` , which reports:

::

    Draw Page shapes:
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 0
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 1
      Shape service: com.sun.star.drawing.ConnectorShape; z-order: 2
      Shape service: com.sun.star.drawing.GroupShape; z-order: 3

The two ellipses have disappeared, replaced by a single GroupShape_.

15.2.2 Binding Shapes
---------------------

Instead of ``_group_ellipses()``, it's possible to call ``_bind_ellipses()`` in |grouper_py|_:

.. tabs::

    .. code-tab:: python

        # Grouper.main() of grouper.py
        s1 = Draw.draw_ellipse(slide=curr_slide, x=x, y=y1, width=width, height=height)
        s2 = Draw.draw_ellipse(slide=curr_slide, x=x, y=y2, width=width, height=height)
        self._bind_ellipses(slide=curr_slide, s1=s1, s2=s2)

The function is defined as:

.. tabs::

    .. code-tab:: python

        # _bind_ellipses() class of grouper.py
        def _bind_ellipses(self, slide: XDrawPage, s1: XShape, s2: XShape) -> None:
            shapes = Lo.create_instance_mcf(
                XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True
            )
            shapes.add(s1)
            shapes.add(s2)
            binder = Lo.qi(XShapeBinder, slide, True)
            binder.bind(shapes)

An empty XShapes_ shape is created, then filled with the component shapes.
The shapes inside XShapes_ are converted into a single object ``XShapeBinder.bind()``.

The result is like the grouped ellipses but with a connector linking the shapes, as in :numref:`ch15fig_bound_ellipses`.

Run the |grouper|_ example with these args.

.. code::

    python -m start -k bind

..
    figure 7

.. cssclass:: screen_shot invert

    .. _ch15fig_bound_ellipses:
    .. figure:: https://user-images.githubusercontent.com/4193389/200202469-230fdae2-34a9-43b9-a7dd-b3b93c7e4096.png
        :alt: The Bound Ellipses.
        :figclass: align-center

        :The Bound Ellipses.

The result is also visible in a call to :py:meth:`.Draw.show_shapes_info`:

::

    Draw Page shapes:
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 0
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 1
      Shape service: com.sun.star.drawing.ConnectorShape; z-order: 2
      Shape service: com.sun.star.drawing.ClosedBezierShape; z-order: 3

The two ellipses have been replaced by a closed Bezier shape.

It's likely easier to link shapes explicitly with connectors, using code like that in ``_connect_rectangles()`` from :ref:`ch15_connecting_tow_rectangles`.
If the result needs to be a single shape, then grouping (not binding) can be applied to the shapes and the connector.

15.2.3 Combining Shapes with XShapeCombiner
-------------------------------------------

|grouper_py|_ calls ``_combine_ellipse()`` to combine the two ellipses:

.. tabs::

    .. code-tab:: python

        # in Grouper.main() of grouper.py
        s1 = Draw.draw_ellipse(slide=curr_slide, x=x, y=y1, width=width, height=height)
        s2 = Draw.draw_ellipse(slide=curr_slide, x=x, y=y2, width=width, height=height)
        self._combine_ellipses(slide=curr_slide, s1=s1, s2=s2)

``_combine_ellipse()`` employs the XShapeCombiner_ interface, which is used in the same way as XShapeBinder_:

.. tabs::

    .. code-tab:: python

        # _combine_ellipses() of grouper.py
        def _combine_ellipses(self, slide: XDrawPage, s1: XShape, s2: XShape) -> None:
            shapes = Lo.create_instance_mcf(
                XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True
            )
            shapes.add(s1)
            shapes.add(s2)
            combiner = Lo.qi(XShapeCombiner, slide, True)
            combiner.combine(shapes)

The combined shape only differs from grouping if the two ellipses are initially overlapping.
:numref:`ch15fig_combining_shape_combiner` shows that the intersecting areas of the two shapes is removed from the combination.

Run the |grouper|_ example with these args.

.. code::

    python -m start -o -k combine

..
    figure 8

.. cssclass:: screen_shot invert

    .. _ch15fig_combining_shape_combiner:
    .. figure:: https://user-images.githubusercontent.com/4193389/200203013-e959dd98-a437-4733-9246-a71226981b74.png
        :alt: Combining Shapes with X-Shape-Combiner.
        :figclass: align-center

        :Combining Shapes with XShapeCombiner_.

The result is also visible in a call to :py:meth:`.Draw.show_shapes_info`:

::

    Draw Page shapes:
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 0
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 1
      Shape service: com.sun.star.drawing.ConnectorShape; z-order: 2
      Shape service: com.sun.star.drawing.ClosedBezierShape; z-order: 3

The two ellipses have again been replaced by a closed Bezier shape .

15.2.4 Richer Shape Combination by Dispatch
-------------------------------------------

The drawback of XShapeCombiner_ that it only supports combination, not merging, subtraction, or intersection.
Those effects had to implemented by using dispatches, as shown in ``_combine_rects()`` in |grouper_py|_:

.. tabs::

    .. code-tab:: python

        # in grouper.py
        def _combine_rects(self, doc: XComponent, slide: XDrawPage) -> XShape:
            print()
            print("Combining rectangles ...")
            r1 = Draw.draw_rectangle(slide=slide, x=50, y=20, width=40, height=20)
            r2 = Draw.draw_rectangle(slide=slide, x=70, y=25, width=40, height=20)
            shapes = Lo.create_instance_mcf(
                XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True
            )
            shapes.add(r1)
            shapes.add(r2)
            comb = Draw.combine_shape(doc=doc, shapes=shapes, combine_op=ShapeCombKind.COMBINE)
            return comb

The dispatching is performed by :py:meth:`.Draw.combine_shape`, which is passed an array of XShapes_ and a constant representing one of the four combining techniques.

:numref:`ch15fig_four_way_shapes` shows the results when the two rectangles created in ``_combine_rects()`` are combined in the different ways.

..
    figure 9

.. cssclass:: screen_shot invert

    .. _ch15fig_four_way_shapes:
    .. figure:: https://user-images.githubusercontent.com/4193389/200207757-228eb5ea-b71a-4e47-98bc-15e5a6e184bf.png
        :alt: The Four Ways of Combining Shapes.
        :figclass: align-center

        :The Four Ways of Combining Shapes.

The merging change in :numref:`ch15fig_four_way_shapes` is a bit subtle – notice that there's no black outline between the rectangles after merging; the merged object is a single shape.

When ``_combine_rects()`` returns, :py:meth:`.Draw.show_shapes_info` reports:

::

    Draw Page shapes:
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 0
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 1
      Shape service: com.sun.star.drawing.ConnectorShape; z-order: 2
      Shape service: com.sun.star.drawing.ClosedBezierShape; z-order: 3
      Shape service: com.sun.star.drawing.PolyPolygonShape; z-order: 4

The combined shape is a PolyPolygonShape_, which means that the shape is created from multiple polygons.

One tricky aspect of combining shapes with dispatches is that the shapes must be selected prior to the dispatch being sent.
After the dispatch has been processed, the selection will have been changed to contain only the single new shape.
This approach is implemented in :py:meth:`.Draw.combine_shape`:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @staticmethod
        def combine_shape(doc: XComponent, shapes: XShapes, combine_op: ShapeCombKind) -> XShape:

            sel_supp = Lo.qi(XSelectionSupplier, GUI.get_current_controller(doc), True)
            sel_supp.select(shapes)

            if combine_op == ShapeCombKind.INTERSECT:
                Lo.dispatch_cmd("Intersect")
            elif combine_op == ShapeCombKind.SUBTRACT:
                Lo.dispatch_cmd("Substract")  # misspelt!
            elif combine_op == ShapeCombKind.COMBINE:
                Lo.dispatch_cmd("Combine")
            else:
                Lo.dispatch_cmd("Merge")

            Lo.delay(500)  # give time for dispatches to arrive and be processed

            # extract the new single shape from the modified selection
            xs = Lo.qi(XShapes, sel_supp.getSelection(), True)
            combined_shape = Lo.qi(XShape, xs.getByIndex(0), True)
            return combined_shape

.. seealso::

    - :py:class:`~.kind.shape_comb_kind.ShapeCombKind`

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`combine_shape`

The shapes are selected by adding them to an XSelectionSupplier_.
The requested dispatch is sent to the selection, and then the function briefly sleeps to ensure that the dispatch has been processed.
An XShapes_ object is obtained from the changed selection, and the new PolyPolygonShape_ is extracted and returned.


15.3 Undoing a Grouping/Binding/Combining
=========================================

Any shapes which have been grouped, bound, or combined can be ungrouped, unbound, or uncombined.
On screen the separated shapes will look the same as before, but may not have the same shape types as the originals.

The ``main()`` function of |grouper_py|_ shows how the combination of the two rectangles can be undone:


.. tabs::

    .. code-tab:: python

        # in Grouper.main() of grouper.py
        # ...
        comp_shape = self._combine_rects(doc=doc, slide=curr_slide)
        # ...
        combiner = Lo.qi(XShapeCombiner, curr_slide, True)
        combiner.split(comp_shape)
        Draw.show_shapes_info(curr_slide)

The combined rectangles shape is passed to ``XShapeCombiner.split()`` which removes the combined shape from the slide, replacing it by its components.
:py:meth:`.Draw.show_shapes_info` shows this result:

::

    Draw Page shapes:
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 0
      Shape service: com.sun.star.drawing.RectangleShape; z-order: 1
      Shape service: com.sun.star.drawing.ConnectorShape; z-order: 2
      Shape service: com.sun.star.drawing.ClosedBezierShape; z-order: 3
      Shape service: com.sun.star.drawing.PolyPolygonShape; z-order: 4
      Shape service: com.sun.star.drawing.PolyPolygonShape; z-order: 5

The last two shapes listed are the separated rectangles, but represented now by two PolyPolygonShape_.

``XShapeCombiner.split()`` only works on shapes that were combined using a ``COMBINE`` dispatch.
Shapes that were composed using merging, subtraction, or intersection, can not be separated.

For grouped and bound shapes, the methods for breaking apart a shape are ``XShapeGrouper.ungroup()`` and ``XShapeBinder.unbind()``.
For example:

.. tabs::

    .. code-tab:: python

        grouper = Lo.qi(XShapeGrouper, curr_slide)
        grouper.ungroup(comp_shape)

15.4 Bezier Curves
==================

The simplest Bezier curve is defined using four coordinates, as in :numref:`ch15fig_cubic_bezier_curve`.

..
    figure 10

.. cssclass:: diagram invert

    .. _ch15fig_cubic_bezier_curve:
    .. figure:: https://user-images.githubusercontent.com/4193389/200210529-d39b8b97-c2e4-4471-bc3a-6ac8787037bd.png
        :alt: A Cubic Bezier Curve.
        :figclass: align-center

        :A Cubic Bezier Curve.

``P0`` and ``P3`` are the start and end points of the curve (also called nodes or anchors), and ``P1`` and ``P2`` are control points,
which specify how the curve bends between the start and finish.
A curve using four points in this way is a cubic Bezier curve, the default type in Office.

The code for generating :numref:`ch15fig_cubic_bezier_curve` is in ``_draw_curve()`` in |draw_bezier_py|_:

.. tabs::

    .. code-tab:: python

        # in bezier_builder.py
        def _draw_curve(self, slide: XDrawPage) -> XShape:
            path_pts: List[Point] = []
            path_flags: List[PolygonFlags] = []

            path_pts.append(Point(1_000, 2_500))
            path_flags.append(PolygonFlags.NORMAL)

            path_pts.append(Point(1_000, 1_000))  # control point
            path_flags.append(PolygonFlags.CONTROL)

            path_pts.append(Point(4_000, 1_000))  # control point
            path_flags.append(PolygonFlags.CONTROL)

            path_pts.append(Point(4_000, 2_500))
            path_flags.append(PolygonFlags.NORMAL)

            return Draw.draw_bezier(slide=slide, pts=path_pts, flags=path_flags, is_open=True)

Most of the curve generation is done by :py:meth:`.Draw.draw_bezier`, but the programmer must still define two list and a boolean.
The ``path_pts[]`` list holds the four coordinates, and ``path_flags[]`` specify their types.
The final boolean argument of :py:meth:`.Draw.draw_bezier` indicates whether the generated curve is to be open or closed.

:numref:`ch15fig_draw_bezier_curve` shows how the curve is rendered.

..
    figure 11

.. cssclass:: diagram invert

    .. _ch15fig_draw_bezier_curve:
    .. figure:: https://user-images.githubusercontent.com/4193389/200211475-d85288a5-380d-4943-8e0e-ee36fd901d61.png
        :alt: The Drawn Bezier Curve
        :figclass: align-center

        :The Drawn Bezier Curve.

:py:meth:`.Draw.draw_bezier` uses the ``is_open`` boolean to decide whether to create an OpenBezierShape_ or a ClosedBezierShape_.
Then it fills a PolyPolygonBezierCoords_ data structure with the coordinates and flags before assigning the structure to the shape's ``PolyPolygonBezier`` property:

.. tabs::

    .. code-tab:: python

        # in the Draw class (simplified)
        @classmethod
        def draw_bezier(
            cls,
            slide: XDrawPage,
            pts: Sequence[Point],
            flags: Sequence[PolygonFlags],
            is_open: bool
        ) -> XShape:

            if len(pts) != len(flags):
                raise IndexError("pts and flags must be the same length")

            bezier_type = "OpenBezierShape" if is_open else "ClosedBezierShape"
            bezier_poly = cls.add_shape(
                slide=slide, shape_type=bezier_type, x=0, y=0, width=0, height=0
            )
            # create space for one bezier shape
            coords = PolyPolygonBezierCoords()
            coords.Coordinates = (pts,)
            coords.Flags = (flags,)

            Props.set(bezier_poly, PolyPolygonBezier=coords)
            return bezier_poly

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`draw_bezier`

A PolyPolygonBezierCoords_ object can store multiple Bezier curves, but :py:meth:`.Draw.draw_bezier` only assigns one curve to it.
Each curve is defined by a list of coordinates and a set of flags.

15.4.1 Drawing a Simple Bezier
------------------------------

The hard part of writing ``_draw_curve()`` in |draw_bezier_py|_ is knowing what coordinates to put into ``path_pts[]``.
Probably the 'easiest' solution is to use a SVG editor to draw the curve by hand, and then extract the coordinates from the generated file.

As the quotes around 'easiest' suggest, this isn't actually that easy since a curve can be much more complex than my example.
A real example may be composed from multiple curves, straight lines, quadratic Bezier sub-curves
(i.e. ones which use only a single control point between anchors), arcs, and smoothing.
The official specification can be found at https://www.w3.org/TR/SVG/paths.html, and there are many tutorials on the topic,
such as https://www.w3schools.com/graphics/svg_path.asp.

Even if you're careful and only draw curves, the generated SVG is not quite the same as the coordinates used by Office's PolyPolygonBezierCoords_.
However, the translation is fairly straightforward, once you've done one or two.

One good online site for drawing simple curves is https://blogs.sitepointstatic.com/examples/tech/svg-curves/cubic-curve.html, developed by Craig Buckler.
It restricts you to manipulating a curve made up of two anchors and two controls, like mine, and displays the corresponding SVG path data, as in :numref:`ch15fig_draw_bezier_curve_ol`.

..
    figure 12

.. cssclass:: diagram invert

    .. _ch15fig_draw_bezier_curve_ol:
    .. figure:: https://user-images.githubusercontent.com/4193389/200213729-9067e159-c8f2-4fc0-ae79-3f13ff0d897b.png
        :alt: The Drawn Bezier Curve
        :figclass: align-center

        :Drawing a Curve Online

:numref:`ch15fig_draw_bezier_curve_ol` is a bit small – the path data at the top-right is:
The path contains two operations: ``M`` and ``C``. ``M`` moves the drawing point to a specified coordinate (in this case (100, 250)).
The ``C`` is followed by three coordinates: (100, 100), (400, 100), and (400, 250).
The first two are the control points and the last is the end point of the curve.
There's no start point since the result of the ``M`` operation is used by default.

Translating this to Office coordinates means using the ``M`` coordinate as the start point,
and applying some scaling of the values to make the curve visible on the page.
Remember that Office uses ``1/100 mm`` units for drawing.
A simple scale factor is to multiply all the numbers by 10, producing: (1000, 2500), (1000, 1000), (4000, 1000), and (4000, 2500).
These are the numbers in :numref:`ch15fig_cubic_bezier_curve`, and utilized by ``_draw_curve()`` in |draw_bezier_py|_.

15.4.2 Drawing a Complicated Bezier Curve
-----------------------------------------

What if you want to draw a curve of more than four points? I use Office's Draw application to draw the curve manually,
save it as an SVG file, and then extract the path coordinates from that file.

Recommend using Draw because it generates path coordinates using ``1/100 mm`` units, which saves me from having to do any scaling.


You might be thinking that if Draw can generate SVG data then why not just import that data as a Bezier curve into the code?
Unfortunately, this isn't quite possible at present – it's true that you can import an SVG file into Office, but it's stored as an image.
In particular, it's available as a GraphicObjectShape_ not a OpenBezierShape_ or a ClosedBezierShape_.
This means that you cannot examine or change its points.

As an example, consider the complex curve in :numref:`ch15fig_draw_bezier_complex_curve` which was created in Draw and exported as an SVG file.

..
    figure 13

.. cssclass:: diagram invert

    .. _ch15fig_draw_bezier_complex_curve:
    .. figure:: https://user-images.githubusercontent.com/4193389/200214680-1618102d-5a26-4053-8603-dbf0c73376a7.png
        :alt: A Complex Bezier Curve, manually produced in Draw.
        :figclass: align-center

        :A Complex Bezier Curve, manually produced in Draw.

Details on how to draw Bezier curves are explained in the Draw user guide, at the end of section 11 on advanced techniques.

The SVG file format is XML-based, so the saved file can be opened by a text editor.

The coordinate information for this OpenBezierShape_ is near the end of the file:

.. code-block:: xml

    <g class="com.sun.star.drawing.OpenBezierShape">
        <g id="id3">
            <path fill="none" stroke="rgb(0,0,0)" d="M 5586,13954 C
            5713,13954 4443,2905 8253,7477 12063,12049 8634,19415 15619,10906
            22604,2397 11682,1381 10285,6334 8888,11287 21207,21447 8253,17002 -
            4701,12557 11174,15986 11174,15986"/>
        </g>
    </g>

The path consists of a single ``M`` operation, and a long ``C`` operation, which should be read as a series of cubic Bezier curves.
Each curve in the ``C`` list is made from three coordinates, since the start point is implicitly given by the initial ``M`` move or the end point of the previous curve in the list.

Copy the data and save it as two lines in a text file (:abbreviation:`e.g.` in ``bpts2.txt``):

::

    M 5586,13954

    C 5713,13954 4443,2905 8253,7477 12063,12049 8634,19415 15619,10906
    22604,2397 11682,1381 10285,6334 8888,11287 21207,21447 8253,17002 -
    4701,12557 11174,15986 11174,15986

Run the |draw_bezier_curve|_ with these args.

.. code::

    python -m start 2

the curve shown in :numref:`ch15fig_bezier_builder_curve` appears on a page.

..
    figure 14

.. cssclass:: diagram invert

    .. _ch15fig_bezier_builder_curve:
    .. figure:: https://user-images.githubusercontent.com/4193389/200215918-6e27d2d2-3195-4550-aa2d-66777718af71.png
        :alt: The Curve Drawn by bezier builder python file
        :figclass: align-center

        :The Curve Drawn by |draw_bezier_py|_

The |draw_bezier_py|_ a data-reading functions can only handle a single ``M`` and ``C`` operation.
If the curve you draw has straight lines, arcs, smoothing, or multiple parts, then the SVG file will contain operations that are not able to be processed by that code.

However, the data-reading functions do recognize the ``Z`` operation, which specifies that the curve should be closed.
If ``Z`` is added as a new line at the end of the ``bpts2.txt``, then the closed Bezier curve in :numref:`ch15fig_bezier_builder_curve_closed` is generated.

..
    figure 15

.. cssclass:: diagram invert

    .. _ch15fig_bezier_builder_curve_closed:
    .. figure:: https://user-images.githubusercontent.com/4193389/200216476-e8552809-f579-4575-9fe9-1f7827059bd1.png
        :alt: The Closed Curve Drawn by bezier builder python file
        :figclass: align-center

        :The Closed Curve Drawn by |draw_bezier_py|_

.. _LineShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineShape.html
.. _ConnectorShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1ConnectorShape.html

.. |grouper| replace:: Grouper
.. _grouper: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_grouper

.. |grouper_py| replace:: grouper.py
.. _grouper_py: https://github.com/Amourspir

.. |draw_bezier_curve| replace:: Draw Bezier Curve
.. _draw_bezier_curve: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_bezier_builder

.. |draw_bezier_py| replace:: bezier_builder.py
.. _draw_bezier_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/draw/odev_bezier_builder/bezier_builder.py

.. _ClosedBezierShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1ClosedBezierShape.html
.. _ConnectorProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1ConnectorProperties.html
.. _ConnectorShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1ConnectorShape.html
.. _GluePoint2: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1drawing_1_1GluePoint2.html
.. _GraphicObjectShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicObjectShape.html
.. _GroupShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GroupShape.html
.. _LineProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1LineProperties.html
.. _OpenBezierShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1OpenBezierShape.html
.. _PolyPolygonBezierCoords: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1drawing_1_1PolyPolygonBezierCoords.html
.. _PolyPolygonShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1PolyPolygonShape.html
.. _PolyPolygonShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1PolyPolygonShape.html
.. _RectangleShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1RectangleShape.html
.. _XGluePointsSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XGluePointsSupplier.html
.. _XSelectionSupplier: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1view_1_1XSelectionSupplier.html
.. _XShapeBinder: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShapeBinder.html
.. _XShapeCombiner: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShapeCombiner.html
.. _XShapeGrouper: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShapeGrouper.html
.. _XShapeGrouper: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShapeGrouper.html
.. _XShapes: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShapes.html
