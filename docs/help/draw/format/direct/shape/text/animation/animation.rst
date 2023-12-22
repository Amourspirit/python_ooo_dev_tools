.. _help_draw_format_direct_shape_text_animation_animations:

Draw Direct Shape Text - Animations
===================================

Shapes can have their text animated. The dialog for this is seen in :ref:`dbdd51ee-038d-4c4c-b097-2105827ef2ea`.

|odev| has several classes for working with the animation of shapes. The classes are:

.. cssclass:: ul-list

    - :py:class:`ooodev.format.draw.direct.text.animation.NoEffect`
    - :py:class:`ooodev.format.draw.direct.text.animation.Blink`
    - :py:class:`ooodev.format.draw.direct.text.animation.ScrollIn`
    - :py:class:`ooodev.format.draw.direct.text.animation.ScrollThrough`
    - :py:class:`ooodev.format.draw.direct.text.animation.ScrollBackForth`


Here is an example of setting the animation of a shape:

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno

        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.direct.text.animation import ScrollBackForth
        from ooodev.format.draw.direct.text.animation import TextAnimationDirection


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = DrawDoc(Draw.create_draw_doc())
                doc.set_visible()
                Lo.delay(700)
                doc.zoom(ZoomKind.ZOOM_75_PERCENT)

                slide = doc.get_slide()

                width = 50
                height = 50
                x = 10
                y = 10

                rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
                cursor = rect.get_shape_text_cursor()
                cursor.append_para("Hello World!")

                txt_scroll_back = ScrollBackForth(
                    start_inside=True,
                    visible_on_exit=True,
                    increment=0.6,
                    delay=350,
                    direction=TextAnimationDirection.UP,
                    count=6,
                )
                txt_scroll_back.apply(rect.component)

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _dbdd51ee-038d-4c4c-b097-2105827ef2ea:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/dbdd51ee-038d-4c4c-b097-2105827ef2ea
        :alt: Shape Text Animation Dialog
        :figclass: align-center

        Shape Text Animation Dialog

Get Shape Text Animation
^^^^^^^^^^^^^^^^^^^^^^^^

We can get the text animation of the shape by using the ``.from_obj()`` method of the animation class.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.text.animation import ScrollBackForth
        # ... other code

        f_style = ScrollBackForth.from_obj(rect.component)
        assert f_style.prop_delay == 350

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.text.animation.NoEffect`
        - :py:class:`ooodev.format.draw.direct.text.animation.Blink`
        - :py:class:`ooodev.format.draw.direct.text.animation.ScrollIn`
        - :py:class:`ooodev.format.draw.direct.text.animation.ScrollThrough`
        - :py:class:`ooodev.format.draw.direct.text.animation.ScrollBackForth`