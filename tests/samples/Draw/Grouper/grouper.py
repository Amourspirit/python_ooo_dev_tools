from __future__ import annotations

import uno
from com.sun.star.container import XNameContainer
from com.sun.star.drawing import XDrawPage
from com.sun.star.drawing import XShape
from com.sun.star.drawing import XShapeBinder
from com.sun.star.drawing import XShapeCombiner
from com.sun.star.drawing import XShapes
from com.sun.star.lang import XComponent

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.office.draw import Draw, ShapeCombKind, DrawingShapeKind, GluePointsKind, GraphicStyleKind, mEx
from ooodev.utils.color import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.info import Info
from ooodev.utils.kind.graphic_arrow_style_kind import GraphicArrowStyleKind


class Grouper:
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

            # create two ellipses
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

            # split the combination into component shapes
            print()
            print("Splitting the combination ...")
            # split the combination into component shapes
            combiner = Lo.qi(XShapeCombiner, curr_slide, True)
            combiner.split(comp_shape)
            Draw.show_shapes_info(curr_slide)

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

        Draw.set_style(shape=conn_shape, graphic_styles=g_styles, style_name=GraphicStyleKind.ARROW_LINE)
        # arrow added at the 'from' end of the connector shape
        # and it thickens line and turns it black

        # use GraphicArrowStyleKind to lookup the values for LineStartName and LineEndName.
        # these are the the same names as seen in Draw, Graphic Sytles: Arrow Line dialog box.
        Props.set(
            conn_shape,
            LineWidth=50,
            FillColor=CommonColor.DARK_BLUE,
            LineStartName=str(GraphicArrowStyleKind.ARROW_SHORT),
            LineStartCenter=False,
            LineEndName=str(GraphicArrowStyleKind.DIAMOND),
        )
        # Props.show_obj_props("Connector Shape", conn_shape)

        # report the glue points for the blue rectangle
        try:
            gps = Draw.get_glue_points(blue_rect)
            print("Glue Points for blue rectangle")
            for i, gp in enumerate(gps):
                print(f"  Glue point {i}: ({gp.Position.X}, {gp.Position.Y})")
        except mEx.DrawError:
            pass

    def _group_ellipses(self, slide: XDrawPage, s1: XShape, s2: XShape) -> None:
        shape_group = Draw.add_shape(slide=slide, shape_type=DrawingShapeKind.GROUP_SHAPE, x=0, y=0, width=0, height=0)
        shapes = Lo.qi(XShapes, shape_group, True)
        # shape_grouper = Lo.qi(XShapeGrouper, slide, True)
        #   XShapeGrouper is deprecated; use GroupShape instead
        shapes.add(s1)
        shapes.add(s2)

    def _bind_ellipses(self, slide: XDrawPage, s1: XShape, s2: XShape) -> None:
        shapes = Lo.create_instance_mcf(XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True)
        shapes.add(s1)
        shapes.add(s2)
        binder = Lo.qi(XShapeBinder, slide, True)
        binder.bind(shapes)

    def _combine_ellipses(self, slide: XDrawPage, s1: XShape, s2: XShape) -> None:
        shapes = Lo.create_instance_mcf(XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True)
        shapes.add(s1)
        shapes.add(s2)
        combiner = Lo.qi(XShapeCombiner, slide, True)
        combiner.combine(shapes)

    def _combine_rects(self, doc: XComponent, slide: XDrawPage) -> XShape:
        print()
        print("Combining rectangles ...")
        r1 = Draw.draw_rectangle(slide=slide, x=50, y=20, width=40, height=20)
        r2 = Draw.draw_rectangle(slide=slide, x=70, y=25, width=40, height=20)
        shapes = Lo.create_instance_mcf(XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True)
        shapes.add(r1)
        shapes.add(r2)
        comb = Draw.combine_shape(doc=doc, shapes=shapes, combine_op=ShapeCombKind.COMBINE)
        return comb
