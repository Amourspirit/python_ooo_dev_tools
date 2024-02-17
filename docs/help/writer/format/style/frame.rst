.. _help_writer_format_style_frame:

Write Style Frame Class
=======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2


Applying Page Frame Styles can be accomplished using the :py:class:`ooodev.format.writer.style.Frame` class.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.writer.style import Frame, StyleFrameKind
        from ooodev.format.writer.modify.frame.area import Color as FrameAreaColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.office.write import Write
        from ooodev.units import UnitMM
        from ooodev.utils.color import StandardColor

        def main() -> int:

            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                cursor = Write.get_cursor(doc)

                txt = "Hello"
                Write.append(cursor=cursor, text=txt)

                style = Frame(name=StyleFrameKind.FRAME)
                tf = Write.add_text_frame(
                    cursor=cursor,
                    ypos=UnitMM(20),
                    text="World",
                    width=UnitMM(40),
                    height=UnitMM(40),
                    styles=[style],
                )

                frm_area_color = FrameAreaColor(
                    color=StandardColor.BRICK_LIGHT2, style_name=StyleFrameKind.FRAME
                )
                frm_area_color.apply(doc)

                f_style = Frame.from_obj(tf)
                assert f_style.prop_name == style.prop_name

                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Apply Style
-----------

In this example we will insert a text frame and apply the ``Frame`` style as seen in :numref:`235807180-852e5112-3eb0-425f-9197-f5d40f9ee2f9` to it.
Then we will apply the area color ``Brick Light 2`` to the ``Frame`` style.

The result of running the following script can be seen in :numref:`235807301-822a616c-7265-4dc6-9e6f-7da7cdaae90e`.

.. tabs::

    .. code-tab:: python

        # ... other code
        txt = "Hello"
        Write.append(cursor=cursor, text=txt)

        style = Frame(name=StyleFrameKind.FRAME)
        # create a frame and apply the frame style to the text frame
        tf = Write.add_text_frame(
            cursor=cursor,
            ypos=UnitMM(20),
            text="World",
            width=UnitMM(40),
            height=UnitMM(40),
            styles=[style],
        )

        # create a frame area color and apply it to the frame style
        frm_area_color = FrameAreaColor(color=StandardColor.BRICK_LIGHT2, style_name=StyleFrameKind.FRAME)
        frm_area_color.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _235807180-852e5112-3eb0-425f-9197-f5d40f9ee2f9:
    .. figure:: https://user-images.githubusercontent.com/4193389/235807180-852e5112-3eb0-425f-9197-f5d40f9ee2f9.png
        :alt: Frame Style
        :figclass: align-center

        Frame Style

    .. _235807301-822a616c-7265-4dc6-9e6f-7da7cdaae90e:
    .. figure:: https://user-images.githubusercontent.com/4193389/235807301-822a616c-7265-4dc6-9e6f-7da7cdaae90e.png
        :alt: Styles applied to Frame Page
        :figclass: align-center
        :width: 550px

        Styles applied to Frame Page

Get Style from Cursor
---------------------

.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = Frame.from_obj(tf)
        assert f_style.prop_name == style.prop_name

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.office.write.Write`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.style.Frame`
