.. _ch16:

*************************
Chapter 16. Making Slides
*************************

.. topic:: Overview

    Creating Slides: title, subtitle, bullets, images, video, buttons; Shape Animations; Dispatch Shapes (special symbols, block arrows, 3D shapes, flowchart elements, callouts, and stars); Slide Viewing

    Examples: |make_slides|_.


The |make_slides|_ example creates a deck of five slides, illustrating different aspects of slide generation:

.. cssclass:: ul-list

    * Slide 1. A slide combining a title and subtitle (see :numref:`ch16fig_title_subtitle`);
    * Slide 2. A slide with a title, bullet points, and an image (see :numref:`ch16fig_slide_title_bullte_img`);
    * Slide 3. A slide with a title, and an embedded video which plays automatically when that slide appears during a slide show (see :numref:`ch16fig_slide_video_frame`);
    * Slide 4. A slide with an ellipse and a rounded rectangle acting as buttons. During a slide show, clicking on the ellipse starts a video playing in an external viewer. Clicking on the rounded rectangle causes the slide show to jump to the first slide in the deck (see :numref:`ch16fig_slide_btns_two`);
    * Slide 5. This slide contains eight shapes generated using dispatches, including special symbols, block arrows, 3D shapes, flowchart elements, callouts, and stars (see :numref:`ch16fig_gui_dispatch_shapes`).

|make_slides_py|_ creates a slide deck, adds the five slides to it, and finishes by asking if you want to close the document.:


.. tabs::

    .. code-tab:: python

        # in make_slides.py
        def main(self) -> None:
            loader = Lo.load_office(Lo.ConnectPipe())

            try:
                doc = Draw.create_impress_doc(loader)
                curr_slide = Draw.get_slide(doc=doc, idx=0)

                GUI.set_visible(is_visible=True, odoc=doc)
                Lo.delay(1_000)  # delay to make sure zoom takes
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                Draw.title_slide(
                    slide=curr_slide, title="Python-Generated Slides", sub_title="Using LibreOffice"
                )

                # second slide
                curr_slide = Draw.add_slide(doc)
                self._do_bullets(curr_slide=curr_slide)

                # third slide: title and video
                curr_slide = Draw.add_slide(doc)
                Draw.title_only_slide(slide=curr_slide, header="Clock Video")
                Draw.draw_media(slide=curr_slide, fnm=self._fnm_clock, x=20, y=70, width=50, height=50)

                # fourth slide
                curr_slide = Draw.add_slide(doc)
                self._button_shapes(curr_slide=curr_slide)

                # fifth slide
                if DrawDispatcher:
                    # windows only
                    # a bit slow due to gui interaction but a good demo
                    self._dispatch_shapes(doc)

                Lo.print(f"Total no. of slides: {Draw.get_slides_count(doc)}")

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

The five slides are explained in the following sections.

16.1 The First Slide (Title and Subtitle)
=========================================

:py:meth:`.Draw.create_impress_doc` calls :py:meth:`.Lo.create_doc`, supplying it with the Impress document string type:

.. tabs::

    .. code-tab:: python

        # in Draw class
        @staticmethod
        def create_impress_doc(loader: XComponentLoader) -> XComponent:
            return Lo.create_doc(doc_type=Lo.DocTypeStr.IMPRESS, loader=loader)

This creates a new slide deck with one slide whose layout depends on Impress' default settings.
:numref:`ch16fig_impress_default_new` shows the usual layout when a user starts Impress.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch16fig_impress_default_new:
    .. figure:: https://user-images.githubusercontent.com/4193389/200931098-a22c8de5-3578-4322-83a3-f1520b8a6988.png
        :alt: The Default New Slide in Impress
        :width: 550px
        :figclass: align-center

        :The Default New Slide in Impress.

The slide contains two empty presentation shapes – the text rectangle at the top is a TitleTextShape_, and the larger rectangle below is a SubTitleShape_.

This first slide, which is at index position ``0`` in the deck, can be referred to by calling :py:meth:`.Draw.get_slide`:

.. tabs::

    .. code-tab:: python

        curr_slide = Draw.get_slide(doc=doc, idx=0)

This is the same method used to get the first page in a Draw document, so we won't go through it again.
The XDrawPage_ object can be examined by calling :py:meth:`.Draw.show_shapes_info` which lists all the shapes (both draw and presentation ones) on the slide:


.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def show_shapes_info(cls, slide: XDrawPage) -> None:
            print("Draw Page shapes:")
            shapes = cls.get_shapes(slide)
            for shape in shapes:
                cls.show_shape_info(shape)

        @classmethod
        def show_shape_info(cls, shape: XShape) -> None:
            print(f"  Shape service: {shape.getShapeType()}; z-order: {cls.get_zorder(shape)}")

        @staticmethod
        def get_zorder(shape: XShape) -> int:
            return int(Props.get(shape, "ZOrder"))

.. seealso::

    .. cssclass:: src-link

        - :odev_src_draw_meth:`show_shapes_info`
        - :odev_src_draw_meth:`show_shape_info`
        - :odev_src_draw_meth:`get_zorder`

:py:meth:`.Draw.show_shapes_info` output for the first slide is:

::

    Draw Page shapes:
      Shape service: com.sun.star.presentation.TitleTextShape; z-order: 0
      Shape service: com.sun.star.presentation.SubtitleShape; z-order: 1

Obviously, the default layout sometimes isn't the one we want.
One solution would be to delete the unnecessary shapes on the slide, then add the shapes that we do want.
A better approach is the programming equivalent of selecting a different slide layout.

This is implemented as several :py:class:`~.draw.Draw` methods, called :py:meth:`.Draw.title_slide`, :py:meth:`.Draw.bullets_slide`, :py:meth:`.Draw.title_only_slide`,
and :py:meth:`.Draw.blank_slide`, which change the slide's layout to those shown in :numref:`ch16fig_slide_layout_methods`.

..
    figure 2

.. cssclass:: screen_shot invert

    .. _ch16fig_slide_layout_methods:
    .. figure:: https://user-images.githubusercontent.com/4193389/200900590-9fe05fc2-c2a1-4d34-8bc8-396e4ed89263.png
        :alt: Slide Layout Methods
        :figclass: align-center

        :Slide Layout Methods.

A title/subtitle layout is used for the first slide by calling:

..
    figure 3

.. cssclass:: screen_shot invert

    .. _ch16fig_title_subtitle:
    .. figure:: https://user-images.githubusercontent.com/4193389/200902224-f9fbdc38-9c69-478a-9b2b-8bf69e3e6257.png
        :alt: The Title and Subtitle Slide.
        :figclass: align-center

        :The Title and Subtitle Slide.

Having a :py:meth:`.Draw.title_slide` method may seem a bit silly since we've seen that the first slide already uses this layout (e.g. in :numref:`ch16fig_impress_default_new`).
That's true for the Impress setup, but may not be the case for other installations with different configurations.

The other layouts shown on the right of :numref:`ch16fig_impress_default_new` could also be implemented as Draw methods, but the four in :numref:`ch16fig_slide_layout_methods` seem most useful.
They set the ``Layout`` property in the DrawPage_ service in the ``com.sun.star.presentation`` module (not the one in the drawing module).

The documentation for DrawPage_ (use ``lodoc DrawPage presentation service``) only says that ``Layout`` stores a short; it doesn't list the possible values or how they correspond to layouts.

For this reason |odev| has :py:class:`~.kind.presentation_layout_kind.PresentationLayoutKind`
which is used as the basis of the layout constants in the :py:class:`~.draw.Draw` class.

:py:meth:`.Draw.title_slide` starts by setting the slide's ``Layout`` property to :py:attr:`.PresentationLayoutKind.TITLE_SUB`:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def title_slide(cls, slide: XDrawPage, title: str, sub_title: str = "") -> None:

            Props.set(slide, Layout=PresentationLayoutKind.TITLE_SUB.value)

            xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.TITLE_TEXT)
            txt_field = Lo.qi(XText, xs, True)
            txt_field.setString(title)

            if sub_title:
                xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.SUBTITLE_TEXT)
                txt_field = Lo.qi(XText, xs, True)
                txt_field.setString(sub_title)

.. seealso::

    .. cssclass:: src-link

        - :odev_src_draw_meth:`title_slide`


This changes the slide's layout to an empty TitleTextShape_ and SubtitleShape_.
The functions adds title and subtitle strings to these shapes, and returns.
The tricky part is obtaining a reference to a particular shape so it can be modified.

One (bad) solution is to use the index ordering of the shapes on the slide, which is displayed by :py:meth:`.Draw.show_shapes_info`.
It turns out that TitleTextShape_ is first (i.e. at index ``0``), and SubtitleShape_ second.
This can be used to write the following code:

.. tabs::

    .. code-tab:: python

        x_shapes = Lo.qi(XShapes, curr_slide)

        title_shape = Lo.qi(XShape, x_shapes.getByIndex(0))
        sub_title_shape = Lo.qi(XShape, x_shapes.getByIndex(1))

This is a bit hacky, so :py:meth:`.Draw.find_shape_by_type` is coded instead, which searches for a shape based on its type:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def find_shape_by_type(cls, slide: XDrawPage, shape_type: DrawingNameSpaceKind | str) -> XShape:

            shapes = cls.get_shapes(slide)
            if not shapes:
                raise ShapeMissingError("No shapes were found in the draw page")

            st = str(shape_type)

            for shape in shapes:
                if st == shape.getShapeType():
                    return shape
            raise ShapeMissingError(f'No shape found for "{st}"')

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`find_shape_by_type`

|odev| has :py:class:`~.kind.drawing_name_space_kind.DrawingNameSpaceKind` to lookup shape type names.

This allows for finding the title shape by calling:

.. tabs::

    .. code-tab:: python

        xs = Draw.find_shape_by_type(curr_slide, DrawingNameSpaceKind.TITLE_TEXT)

16.2 The Second Slide (Title, Bullets, and Image)
=================================================

The second slide uses a title and bullet points layout, with an image added at the bottom right corner. The relevant lines in |make_slides_py|_ are:

.. tabs::

    .. code-tab:: python

        # in main() in make_slides.py
        curr_slide = Draw.add_Slide(doc)
        self._do_bullets(curr_slide=curr_slide)

The result shown in :numref:`ch16fig_slide_title_bullte_img`.

..
    figure 4

.. cssclass:: screen_shot invert

    .. _ch16fig_slide_title_bullte_img:
    .. figure:: https://user-images.githubusercontent.com/4193389/200941913-ef233dc5-b14b-4ca8-a3e7-640c64e90fdf.png
        :alt: A Slide with a Title, Bullet Points, and an Image.
        :width: 525px
        :figclass: align-center

        :A Slide with a Title, Bullet Points, and an Image.

:numref:`ch16fig_slide_title_bullte_img` slide is created by ``_do_bullets()`` in |make_slides_py|_:

.. tabs::

    .. code-tab:: python

        # in main() in make_slides.py
        def _do_bullets(self, curr_slide: XDrawPage) -> None:
            # second slide: bullets and image
            body = Draw.bullets_slide(slide=curr_slide, title="What is an Algorithm?")

            # bullet levels are 0, 1, 2,...
            Draw.add_bullet(
                bulls_txt=body,
                level=0,
                text="An algorithm is a finite set of unambiguous instructions for solving a problem.",
            )

            Draw.add_bullet(
                bulls_txt=body,
                level=1,
                text=("An algorithm is correct if on all legitimate inputs,",
                    " it outputs the right answer in a finite amount of time"),
            )

            Draw.add_bullet(bulls_txt=body, level=0, text="Can be expressed as")
            Draw.add_bullet(bulls_txt=body, level=1, text="pseudocode")
            Draw.add_bullet(bulls_txt=body, level=0, text="flow charts")
            Draw.add_bullet(bulls_txt=body, level=1, text="text in a natural language (e.g. English)")
            Draw.add_bullet(bulls_txt=body, level=1, text="computer code")
            # add the image in bottom right corner, and scaled if necessary
            im = Draw.draw_image_offset(
                slide=curr_slide, fnm=self._fnm_img, xoffset=ImageOffset(0.6), yoffset=ImageOffset(0.5)
            )
            # move below the slide text
            Draw.move_to_bottom(slide=curr_slide, shape=im)

:py:meth:`.Draw.bullets_slide` works in a similar way to :py:meth:`.Draw.title_slide` – first the slide's layout is set, then the presentation shapes are found and modified:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def bullets_slide(cls, slide: XDrawPage, title: str) -> XText:

            Props.set(slide, Layout=PresentationLayoutKind.TITLE_BULLETS.value)

            xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.TITLE_TEXT)
            txt_field = Lo.qi(XText, xs, True)
            txt_field.setString(title)

            xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.BULLETS_TEXT)
            return Lo.qi(XText, xs, True)

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`bullets_slide`

The :py:attr:`.PresentationLayoutKind.TITLE_BULLETS` enum changes the slide's layout to contain two presentation shapes – a TitleTextShape_ at the top,
and an OutlinerShape_ beneath it (as in the second picture in :numref:`ch16fig_slide_layout_methods`).
:py:meth:`.Draw.bullets_slide` calls :py:meth:`.Draw.find_shape_by_type` twice to find these shapes, but it does nothing to the OutlinerShape_ itself,
returning it as an XText_ reference. This allows text to be inserted into the shape by other code (i.e. by :py:meth:`.Draw.add_bullet`).


16.2.1 Adding Bullets to a Text Area
------------------------------------

:py:meth:`.Draw.add_bullet` converts the shape's XText_ reference into an XTextRange_, which offers a ``setString()`` method:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @staticmethod
        def add_bullet(bulls_txt: XText, level: int, text: str) -> None:

            bulls_txt_end = Lo.qi(XTextRange, bulls_txt, True).getEnd()
            Props.set(bulls_txt_end, NumberingLevel=level)
            bulls_txt_end.setString(f"{text}\n")

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`add_bullet`

As explained :ref:`ch05`, XTextRange_ is part of the TextRange_ service which inherits both paragraph and character property classes, as indicated by :numref:`ch16fig_text_rng_service`.

..
    figure 5

.. cssclass:: diagram invert

    .. _ch16fig_text_rng_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/200949420-c011120a-9cb9-43d6-aa0d-87a3377d5ceb.png
        :alt: The Text Range Service.
        :figclass: align-center

        :The TextRange_ Service.

A look through the ParagraphProperties_ documentation reveals a ``NumberingLevel`` property which affects the displayed bullet level.

Another way of finding out about the properties associated with XTextRange_ is to use :py:meth:`.Props.show_obj_props` to list all of them:

.. tabs::

    .. code-tab:: python

        Props.show_obj_props("TextRange in OutlinerShape", tr)

The bullet text is added with ``XTextRange.setString()``.
A newline is added to the text before the set, to ensure that the string is treated as a complete paragraph.
The drawback is that the newline causes an extra bullet symbol to be drawn after the real bullet points.
This can be seen in :numref:`ch16fig_slide_title_bullte_img`, at the bottom of the slide. (Principal Skinner is pointing at it.)

16.2.2 Offsetting an Image
--------------------------

The |animate_bike|_ example in :ref:`ch14` employed a version of :py:meth:`.Draw.draw_image` based around specifying an (x, y) position on the page and a width and height for the image frame.
:py:meth:`.Draw.draw_image_offset` used here is a variant which specifies its position in terms of fractional offsets from the top-left corner of the slide.

.. tabs::

    .. code-tab:: python

        from ooodev.office.draw import Draw, ImageOffset

        im = Draw.draw_image_offset(
            slide=curr_slide, fnm="skinner.png", xoffset=ImageOffset(0.6), yoffset=ImageOffset(0.5)
        )

The last two arguments mean that the image's top-left corner will be placed at a point that is ``0.6`` of the slide's width across and ``0.5`` of its height down.
:py:meth:`~.Draw.draw_image_offset` also scales the image so that it doesn't extend beyond the right and bottom edges of the slide.
The scaling is the same along both dimensions so the picture isn't distorted.

:py:class:`~.image_offset.ImageOffset` ensure that offsets are not out of range.

The code for :py:meth:`.Draw.draw_image_offset`:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def draw_image_offset(
            cls, slide: XDrawPage, fnm: PathOrStr, xoffset: ImageOffset, yoffset: ImageOffset
        ) -> XShape:

            slide_size = cls.get_slide_size(slide)
            x = round(slide_size.Width * xoffset.Value)  # in mm units
            y = round(slide_size.Height * yoffset.Value)

            max_width = slide_size.Width - x
            max_height = slide_size.Height - y

            im_size = ImagesLo.calc_scale(fnm=fnm, max_width=max_width, max_height=max_height)
            if im_size is None:
                Lo.print(f'Unalbe to calc image size for "{fnm}"')
                return None
            return cls.draw_image(
                slide=slide, fnm=fnm, x=x, y=y, width=im_size.Width, height=im_size.Height
            )

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`draw_image_offset`

:py:meth:`~.Draw.draw_image_offset` uses the slide's size to determine an (x, y) position for the image, and its width and height.
:py:meth:`.ImagesLo.calc_scale` calculates the best width and height for the image frame such that the image will be drawn entirely on the slide:

.. tabs::

    .. code-tab:: python

        # in ImagesLo class
        @classmethod
        def calc_scale(cls, fnm: PathOrStr, max_width: int, max_height: int) -> Size | None:
            im_size = cls.get_size_100mm(fnm)  # in 1/100 mm units
            if im_size is None:
                return None

            width_scale = (max_width * 100) / im_size.Width
            height_scale = (max_height * 100) / im_size.Height

            scale_factor = min(width_scale, height_scale)

            w = round(im_size.Width * scale_factor / 100)
            h = round(im_size.Height * scale_factor / 100)
            return Size(w, h)

:py:meth:`~.ImagesLo.calc_scale` uses :py:meth:`.ImagesLo.get_size100mm` to retrieve the size of the image in ``1/100 mm`` units, and then a scale factor is calculated for both the width and height.
This is used to set the image frame's dimensions when the graphic is loaded by :py:meth:`~.Draw.draw_image`.

16.3 The Third Slide (Title and Video)
======================================

The third slide consists of a title shape and a video frame, which looks like :numref:`ch16fig_slide_video_frame`.

..
    figure 6

.. cssclass:: screen_shot invert

    .. _ch16fig_slide_video_frame:
    .. figure:: https://user-images.githubusercontent.com/4193389/200954466-2b1e2176-1835-4f54-bee0-4888c090d5c1.png
        :alt: A Slide Containing a Video Frame.
        :figclass: align-center

        :A Slide Containing a Video Frame.

When this slide appears in a slide show, the video will automatically start playing.

The code for generating this slide is:

.. tabs::

    .. code-tab:: python

        # in MakeSlide.main() of make_slides.py
        curr_slide = Draw.add_slide(doc)
        Draw.title_only_slide(slide=curr_slide, header="Clock Video")
        Draw.draw_media(slide=curr_slide, fnm=self._fnm_clock, x=20, y=70, width=50, height=50)

:py:meth:`.Draw.title_only_slide` works in a similar way to :py:meth:`~.title_slide` and :py:meth:`~.bullets_slide`:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def title_only_slide(cls, slide: XDrawPage, header: str) -> None:

            Props.set(slide, Layout=PresentationLayoutKind.TITLE_ONLY.value)

            xs = cls.find_shape_by_type(slide=slide, shape_type=DrawingNameSpaceKind.TITLE_TEXT)
            txt_field = Lo.qi(XText, xs, True)
            txt_field.setString(header)

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`title_only_slide`

The ``MediaShape`` service doesn't appear in the Office documentation.
Perhaps one reason for its absence is that the shape behaves a little 'erratically'.
Although |make_slides_py|_ successfully builds a slide deck containing the video.
When the deck is run as a slide show, the video frame is sometimes incorrectly placed, although the video plays correctly.

:py:meth:`.Draw.draw_media` is defined as:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def draw_media(
            cls, slide: XDrawPage, fnm: PathOrStr, x: int, y: int, width: int, height: int
        ) -> XShape:

            shape = cls.add_shape(
                slide=slide, shape_type=DrawingShapeKind.MEDIA_SHAPE, x=x, y=y, width=width, height=height
            )

            Lo.print(f'Loading media: "{fnm}"')
            cls.set_shape_props(shape, Loop=True, MediaURL=mFileIO.FileIO.fnm_to_url(fnm))

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`draw_media`

In the absence of documentation, :py:meth:`.Props.show_obj_props` can be used to list the properties for the ``MediaShape``:

.. tabs::

    .. code-tab:: python

        Props.show_obj_props("Shape", shape)

The ``MediaURL`` property requires a file in URL format, and ``Loop`` is a boolean for making the animation play repeatedly.

16.4 The Fourth Slide (Title and Buttons)
=========================================

The fourth slide has two 'buttons' – an ellipse which starts a video playing in an external application, and a rounded rectangle which makes the presentation jump to the first slide.
These actions are both implemented using the ``OnClick`` property for presentation shapes.
:numref:`ch16fig_slide_btns_two` shows how the slide looks.

..
    figure 7

.. cssclass:: screen_shot invert

    .. _ch16fig_slide_btns_two:
    .. figure:: https://user-images.githubusercontent.com/4193389/200957116-abb24fc3-d0e3-4da2-a442-7a0c974a4cca.png
        :alt: A Slide with Two Buttons
        :width: 525px
        :figclass: align-center

        :A Slide with Two 'Buttons'.

The relevant code in ``main()`` of |make_slides_py|_ is:

.. tabs::

    .. code-tab:: python

        curr_slide = Draw.add_slide(doc)
        self._button_shapes(curr_slide=curr_slide)

This button approach to playing a video doesn't suffer from the strange behavior when using ``MediaShape`` on the third slide.

The ``_button_shapes()`` method in |make_slides_py|_ creates the slide:

.. tabs::

    .. code-tab:: python

        def _button_shapes(self, curr_slide: XDrawPage) -> None:
            Draw.title_only_slide(slide=curr_slide, header="Wildlife Video Via Button")

            sz = Draw.get_slide_size(curr_slide)
            width = 80
            height = 40

            ellipse = Draw.draw_ellipse(
                slide=curr_slide,
                x=round((sz.Width - width) / 2),
                y=round((sz.Height - height) / 2),
                width=width,
                height=height,
            )

            Draw.add_text(shape=ellipse, msg="Start Video", font_size=30)
            Props.set(
                ellipse, OnClick=ClickAction.DOCUMENT, Bookmark=FileIO.fnm_to_url(self._fnm_wildlife)
            )
            Props.set(
                ellipse, Effect=AnimationEffect.FADE_FROM_BOTTOM, Speed=AnimationSpeed.SLOW
            )

            # draw a rounded rectangle with text
            button = Draw.draw_rectangle(
                slide=curr_slide, x=sz.Width-width-4, y=sz.Height-height-5, width=width, height=height
            )
            Draw.add_text(shape=button, msg="Click to go\nto slide 1")
            Draw.set_gradient_color(shape=button, name=DrawingGradientKind.SUNSHINE)
            # clicking makes the presentation jump to first slide
            Props.set(button, CornerRadius=300, OnClick=ClickAction.FIRSTPAGE)

A minor point of interest is that a rounded rectangle is a RectangleShape_, but with its ``CornerRadius`` property set.

The more important part of the method is the two uses of the ``OnClick`` property from the presentation Shape class.

Clicking on the ellipse executes the video file that was passed into the constructor of ``MakeSlides`` in |make_slides_py|_.
This requires ``OnClick`` to be assigned the ``ClickAction.DOCUMENT`` constant, and ``Bookmark`` to refer to the file as an URL.

Clicking on the rounded rectangle causes the slide show to jump back to the first page.
This needs ``OnClick`` to be set to ``ClickAction.FIRSTPAGE``.

Several other forms of click action are listed in :numref:`ch16tbl_click_action_effects`.

..
    Table 1

.. _ch16tbl_click_action_effects:

.. table:: ClickAction Effects.
    :name: ClickAction_Effects

    ============== ==========================================================================================
     ClickAction    Name Effect                                                                              
    ============== ==========================================================================================
     NONE           No action is performed on the click. Animation and fade effects are also switched off.   
     PREVPAGE       The presentation jumps to the previous page.                                             
     NEXTPAGE       The presentation jumps to the next page.                                                 
     FIRSTPAGE      The presentation continues with the first page.                                          
     LASTPAGE       The presentation continues with the last page.                                           
     BOOKMARK       The presentation jumps to a bookmark.                                                    
     DOCUMENT       The presentation jumps to another document.                                              
     INVISIBLE      The object renders itself invisible after a click.                                       
     SOUND          A sound is played after a click.                                                         
     VERB           An OLE verb is performed on this object.                                                 
     VANISH         The object vanishes with its effect.                                                     
     PROGRAM        Another program is executed after a click.                                               
     MACRO          An Office macro is executed after the click.                                             
    ============== ==========================================================================================

:numref:`ch16tbl_click_action_effects` shows that it's possible to jump to various places in a slide show, and also execute macros and external programs.
In both cases, the ``Bookmark`` property is used to specify the URL of the macro or program.
For example, the following will invoke Windows' calculator when the button is pressed:

.. tabs::

    .. code-tab:: python

        Props.set(
            button,
            OnClick=ClickAction.PROGRAM,
            Bookmark=FileIO.fnm_to_url(f'(System.getenv("SystemRoot")}\\System32\\calc.exe')
            )

``Bookmark`` requires an absolute path to the application, converted to URL form.

Clicking on the ClickAction_ takes you to a table very like the one in :numref:`ch16tbl_click_action_effects`.

16.5 Shape Animation
====================

Shape animations are performed during a slide show, and are regulated through three presentation Shape properties:
``Effect``, ``Speed`` and ``TextEffect``.

``Effect`` can be assigned a large range of animation effects, which are defined as constants in the AnimationEffect_ enumeration.

Details can be found in the |star_presentation|_ module.
Another nice summary, in the form of a large table, is `in the Developer's Guide <https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Animations_and_Interactions>`_.
:numref:`ch16fig_animation_effect_dev_guide` shows part of that table.

..
    figure 8

.. cssclass:: screen_shot invert

    .. _ch16fig_animation_effect_dev_guide:
    .. figure:: https://user-images.githubusercontent.com/4193389/200963820-001b7e97-c835-4002-83e2-273316d2f9b4.png
        :alt: Animation Effect Constants Table in the Developer's Guide.
        :width: 525px
        :figclass: align-center

        :AnimationEffect_ Constants `Table in the Developer's Guide <https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Animations_and_Interactions>`_.

There are two broad groups of effects: those that move a shape onto the slide when the page appears, and fade effects that make a shape gradually appear in a given spot.

The following code fragment makes the ellipse on the fourth slide slide into view, starting from the left of the slide:

.. tabs::

    .. code-tab:: python

        # in _button_shapes() in make_slides.py
        Props.set(
            ellipse, Effect=AnimationEffect.MOVE_FROM_LEFT, Speed=AnimationSpeed.FAST
        )

The animation speed takes a AnimationSpeed_ value and can be set to  ``AnimationSpeed.SLOW``, ``AnimationSpeed.MEDIUM``, or ``AnimationSpeed.FAST``.

Unfortunately, there seems to be an issue with some of the Animation Effects as shown in :numref:`ch16fig_animationeffect_fade_from_lowerright_bug`,
:numref:`ch16fig_animationeffect_fade_from_bottom_bug`, and :numref:`ch16fig_animationeffect_fade_from_top_dev_tool_view`.
When some of the effects are set they actually work in reverse. At least this is the case on Windows 10 and LibreOffice 7.3
There seemed to be issues with most of the fade effects. Not all effects were tested due to the volume of effects.
There may be more effects of different types not working correctly.

.. cssclass:: screen_shot invert

    .. _ch16fig_animationeffect_fade_from_lowerright_bug:
    .. figure:: https://user-images.githubusercontent.com/4193389/201223650-ed3e195f-f506-4fc3-af5d-a14ea02008bc.png
        :alt: :Animation Effect FADE FROM LOWER RIGHT workS in reverse
        :width: 550px
        :figclass: align-center

        :``AnimationEffect.FADE_FROM_LOWERRIGHT`` reversed


.. cssclass:: screen_shot invert

    .. _ch16fig_animationeffect_fade_from_bottom_bug:
    .. figure:: https://user-images.githubusercontent.com/4193389/201224471-98b499b5-c283-48aa-b8cd-0f4eb5321922.png
        :alt: :Animation Effect FADE FROM BOTTOM workS in reverse
        :width: 550px
        :figclass: align-center

        :``AnimationEffect.FADE_FROM_BOTTOM`` reversed

The developer tools of LibreOffice can be used to confirm that ``Effect`` property is actually being set correctly as shown in :numref:`ch16fig_animationeffect_fade_from_top_dev_tool_view`.
Developer tools are available in LibreOffice ``7.3 +``.

.. cssclass:: screen_shot invert

    .. _ch16fig_animationeffect_fade_from_top_dev_tool_view:
    .. figure:: https://user-images.githubusercontent.com/4193389/201225731-ae40e251-0a13-4eda-8e37-4f2a4a0ee4ad.png
        :alt: :Animation Effect FADE FROM TOP workS in reverse, developer tools view
        :width: 680px
        :figclass: align-center

        :``AnimationEffect.FADE_TOP_BOTTOM`` reversed developer tool view.

More Complex Shape Animations
-----------------------------

If you browse chapter 9 of the Impress user's guide on slide shows, its animation capabilities extend well beyond the constants in ``AnimationEffect``.
These features are available through the XAnimationNode_ interface, which is obtained like so:

.. tabs::

    .. code-tab:: python

        from com.sun.star.animations import XAnimationNode
        from ooodev.utils.lo import Lo

        node_supp = Lo.qi(XAnimationNodeSupplier, slide)
        slide_node = node_supp.getAnimationNode()  # XAnimationNode

XAnimationNode_ allows a programmer much finer control over animation timings and animation paths for shapes.
XAnimationNode_ is part of the large ``com.sun.star.animations`` package.

16.6 The Fifth Slide (Various Dispatch Shapes)
==============================================

The fifth slide is a hacky, slow solution for generating the numerous shapes in Impress' GUI which have no corresponding classes in the API.
The approach uses dispatch commands, |odevgui_win|_, and :external+odevguiwin:ref:`class_robot_keys` (first described back in :ref:`ch04_robot_keys`).

The resulting slide is shown in :numref:`ch16fig_gui_dispatch_shapes`.

..
    figure 9

.. cssclass:: screen_shot

    .. _ch16fig_gui_dispatch_shapes:
    .. figure:: https://user-images.githubusercontent.com/4193389/201233121-e867d84c-cf75-4112-8845-25d3ddbdd64d.png
        :alt: Shapes Created by Dispatch Commands.
        :width: 525px
        :figclass: align-center

        :Shapes Created by Dispatch Commands.

The shapes in :numref:`ch16fig_gui_dispatch_shapes` are just a few of the many available via Impress' "Drawing Toolbar", shown in :numref:`ch16fig_gui_toolbar_shapes`.
The relevant menus are labeled and their sub-menus are shown beneath the toolbar.

..
    figure 10

.. cssclass:: diagram invert

    .. _ch16fig_gui_toolbar_shapes:
    .. figure:: https://user-images.githubusercontent.com/4193389/201736250-445586e0-1e60-48d6-9c13-4548d843a50c.png
        :alt: The Shapes Available from the Drawing Toolbar
        :figclass: align-center

        :The Shapes Available from the Drawing Toolbar.

Each sub-menu shape has a name which appears in a tooltip when the cursor is placed over the shape's icon.
This text turns out to be very useful when writing the dispatch commands.

There's also a "3D-Objects" toolbar which offers the shapes in :numref:`ch16fig_gui_toolbar_3d_objects`.

..
    figure 11

.. cssclass:: diagram invert

    .. _ch16fig_gui_toolbar_3d_objects:
    .. figure:: https://user-images.githubusercontent.com/4193389/201736895-64f54480-4830-4ab8-94ca-bc7701f49fe0.png
        :alt: The 3D Objects Toolbar
        :figclass: align-center

        :The 3D-Objects Toolbar.

Some of these 3D shapes are available in the API as undocumented Shape subclasses, but it was unable to programmatically resize the shapes to make them visible.
The only way possible to get them to appear at a reasonable size was by creating them with dispatch commands.

Although there's no mention of these custom and 3D shapes in the Developer's Guide, their dispatch commands do appear in the
``UICommands.ods`` spreadsheet (available from https://arielch.fedorapeople.org/devel/ooo/UICommands.ods).
They're also mentioned, in less detail, in the online documentation for Impress dispatches at
https://wiki.documentfoundation.org/Development/DispatchCommands#Impress_slots_.28sdslots.29


It's quite easy to match up the tooltip names in the GUI with the dispatch names.
For example, the smiley face in the Symbol shapes menu is called "Smiley Face" in the GUI and ``.uno:SymbolShapes.smiley`` in the ``UICommands`` spreadsheet.

|make_slides_py|_ generates the eight shapes shown in :numref:`ch16fig_gui_dispatch_shapes` by calling ``_dispatch_shapes()``:

.. tabs::

    .. code-tab:: python

        # in make_slides.py
        def _dispatch_shapes(self, doc: XComponent) -> None:
            curr_slide = Draw.add_slide(doc)
            Draw.title_only_slide(slide=curr_slide, header="Dispatched Shapes")

            GUI.set_visible(is_visible=True, odoc=doc)
            Lo.delay(1_000)

            Draw.goto_page(doc=doc, page=curr_slide)
            Lo.print(f"Viewing Slide number: {Draw.get_slide_number(Draw.get_viewed_page(doc))}")

            # first row
            y = 38
            _ = Draw.add_dispatch_shape(
                slide=curr_slide,
                shape_dispatch=ShapeDispatchKind.BASIC_SHAPES_DIAMOND,
                x=20,
                y=y,
                width=50,
                height=30,
                fn=DrawDispatcher.create_dispatch_shape,
            )
            _ = Draw.add_dispatch_shape(
                slide=curr_slide,
                shape_dispatch=ShapeDispatchKind.THREE_D_HALF_SPHERE,
                x=80,
                y=y,
                width=50,
                height=30,
                fn=DrawDispatcher.create_dispatch_shape,
            )
            dshape = Draw.add_dispatch_shape(
                slide=curr_slide,
                shape_dispatch=ShapeDispatchKind.CALLOUT_SHAPES_CLOUD_CALLOUT,
                x=140,
                y=y,
                width=50,
                height=30,
                fn=DrawDispatcher.create_dispatch_shape,
            )
            Draw.set_bitmap_color(shape=dshape, name=DrawingBitmapKind.LITTLE_CLOUDS)

            dshape = Draw.add_dispatch_shape(
                slide=curr_slide,
                shape_dispatch=ShapeDispatchKind.FLOW_CHART_SHAPES_FLOWCHART_CARD,
                x=200,
                y=y,
                width=50,
                height=30,
                fn=DrawDispatcher.create_dispatch_shape,
            )
            Draw.set_hatch_color(shape=dshape, name=DrawingHatchingKind.BLUE_NEG_45_DEGREES)
            # convert blue to black manually
            dhatch = cast(Hatch, Props.get(dshape, "FillHatch"))
            dhatch.Color = CommonColor.BLACK
            Props.set(dshape, LineColor=CommonColor.BLACK, FillHatch=dhatch)
            # Props.show_obj_props("Hatch Shape", dshape)

            # second row
            y = 100
            dshape = Draw.add_dispatch_shape(
                slide=curr_slide,
                shape_dispatch=ShapeDispatchKind.STAR_SHAPES_STAR_12,
                x=20,
                y=y,
                width=40,
                height=40,
                fn=DrawDispatcher.create_dispatch_shape,
            )
            Draw.set_gradient_color(shape=dshape, name=DrawingGradientKind.SUNSHINE)
            Props.set(dshape, LineStyle=LineStyle.NONE)

            dshape = Draw.add_dispatch_shape(
                slide=curr_slide,
                shape_dispatch=ShapeDispatchKind.SYMBOL_SHAPES_HEART,
                x=80,
                y=y,
                width=40,
                height=40,
                fn=DrawDispatcher.create_dispatch_shape,
            )
            Props.set(dshape, FillColor=CommonColor.RED)

            _ = Draw.add_dispatch_shape(
                slide=curr_slide,
                shape_dispatch=ShapeDispatchKind.ARROW_SHAPES_LEFT_RIGHT_ARROW,
                x=140,
                y=y,
                width=50,
                height=30,
                fn=DrawDispatcher.create_dispatch_shape,
            )
            dshape = Draw.add_dispatch_shape(
                slide=curr_slide,
                shape_dispatch=ShapeDispatchKind.THREE_D_CYRAMID,
                x=200,
                y=y - 20,
                width=50,
                height=50,
                fn=DrawDispatcher.create_dispatch_shape,
            )
            Draw.set_bitmap_color(shape=dshape, name=DrawingBitmapKind.STONE)

            Draw.show_shapes_info(curr_slide)

A title-only slide is created, followed by eight calls to :py:meth:`.Draw.add_dispatch_shape` to create two rows of four shapes in :numref:`ch16fig_gui_dispatch_shapes`.

Note that :py:meth:`.Draw.add_dispatch_shape` take a ``fn`` parameter. This is basically a call back function.
``fn`` is expected to be a function that takes a XDrawPage_ and ``str`` as input parameters and returns XShape_ or ``None``.

The reason for this is |odev| is not responsible for automating Windows GUI however, |odevgui_win|_ is.
|odevgui_win|_ provides :external+odevguiwin:py:meth:`odevgui_win.draw_dispatcher.DrawDispatcher.create_dispatch_shape` that handles automating mouse movements and returns the shape.
So, :py:meth:`~.Draw.add_dispatch_shape` is passed as call back function.


.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`add_dispatch_shape`

16.6.1 Viewing the Fifth Slide
------------------------------

:py:meth:`.Draw.add_dispatch_shape` requires the fifth slide to be the active, visible window on- screen.
This necessitates a call to :py:meth:`.GUI.set_visible` to make the document visible, but that isn't quite enough.
Making the document visible causes the first slide to be displayed, not the fifth one.

Impress offers many ways of viewing slides, which are implemented in the API as view classes that inherit the Controller service. The inheritance structure is shown in :numref:`ch16fig_impress_view_classes`.

..
    figure 12

.. cssclass:: diagram invert

    .. _ch16fig_impress_view_classes:
    .. figure:: https://user-images.githubusercontent.com/4193389/201742799-ad85319f-bff0-46d7-9f70-59f4106c16b4.png
        :alt: Impress View Classes.
        :figclass: align-center

        :Impress View Classes.

When a Draw or Impress document is being edited, the view is DrawingDocumentDrawView, which supports a number of useful properties,
such as ``ZoomType`` and ``VisibleArea``. Its XDrawView_ interface is employed for getting and setting the current page displayed in this view.

:py:meth:`.Draw.goto_page` gets the XController_ interface for the document, and converts it to XDrawView_ so the visible page can be set:

.. tabs::

    .. code-tab:: python

        # in Draw class (simplified)
        @classmethod
        def goto_page(cls, doc: XComponent, page: XDrawPage) -> None:
            try:
                ctl = GUI.get_current_controller(doc)
                cls.goto_page(ctl, page)
            except DrawError:
                raise
            except Exception as e:
                raise DrawError("Error while trying to go to page") from e

        @staticmethod
        def goto_page(ctl: XController, page: XDrawPage) -> None:
            try:
                xdraw_view = Lo.qi(XDrawView, ctl)
                xdraw_view.setCurrentPage(page)
            except Exception as e:
                raise DrawError("Error while trying to go to page") from e

.. seealso::

    .. cssclass:: src-link

        :odev_src_draw_meth:`goto_page`

After the call to :py:meth:`.Draw.goto_page`, the specified draw page will be visible on-screen, and so receive any dispatch commands.

:py:meth:`.Draw.get_viewed_page` returns a reference to the currently viewed page by calling ``XDrawView.getCurrentPage()``:

.. tabs::

    .. code-tab:: python

        # in Draw class
        @staticmethod
        def get_viewed_page(doc: XComponent) -> XDrawPage:
            try:
                ctl = GUI.get_current_controller(doc)
                xdraw_view = Lo.qi(XDrawView, ctl, True)
                return xdraw_view.getCurrentPage()
            except Exception as e:
                raise DrawPageError("Error geting Viewed page") from e


16.6.2 Adding a Dispatch Shape to the Visible Page
--------------------------------------------------

If you try adding a smiley face to a slide inside Impress, it's a two-step process.
It isn't enough only to click on the icon, it's also necessary to drag the cursor over the page in order for the shape to appear and be resized.

These steps are necessary for all the Drawing toolbar and 3D-Objects shapes, and are emulated by my code.
The programming equivalent of clicking on the icon is done by calling :py:meth:`.Lo.dispatch_cmd`,
while implementing a mouse drag utilizes |odevgui_win|_ and :external+odevguiwin:ref:`class_robot_keys`.

:py:meth:`Draw.add_dispatch_shape` uses :py:meth:`Draw.create_dispatch_shape` to create the shape, and then positions and resizes it:

.. tabs::

    .. code-tab:: python

        # in Draw class
        @classmethod
        def add_dispatch_shape(
            cls, slide: XDrawPage, shape_dispatch: ShapeDispatchKind | str,
            x: int, y: int, width: int, height: int, fn: DispatchShape
        ) -> XShape:
            cls.warns_position(slide, x, y)
            try:
                shape = fn(slide, str(shape_dispatch))
                if shape is None:
                    raise NoneError(f'Failed to add shape for dispatch command "{shape_dispatch}"')
                cls.set_position(shape=shape, x=x, y=y)
                cls.set_size(shape=shape, width=width, height=height)
                return shape
            except NoneError:
                raise
            except Exception as e:
                raise ShapeError(
                    f'Error occured adding dispatch shape for dispatch command "{shape_dispatch}"'
                ) from e

.. |animate_bike| replace:: Animate Bike
.. _animate_bike: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_animate_bike

.. |make_slides| replace:: Make Slides
.. _make_slides: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_make_slides

.. |make_slides_py| replace:: make_slides.py
.. _make_slides_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/impress/odev_make_slides/make_slides.py

.. |star_presentation| replace:: com.sun.star.presentation
.. _star_presentation: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1presentation.html

.. _AnimationEffect: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1presentation.html#a10f2a3114ab31c0e6f7dc48f656fd260
.. _AnimationSpeed: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1presentation.html#a07b64dc4a366b20ad5052f974ffdbf62
.. _ClickAction: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1presentation.html#a85fe75121d351785616b75b2c5661d8f
.. _DrawPage: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1DrawPage.html
.. _OutlinerShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1OutlinerShape.html
.. _ParagraphProperties: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html
.. _RectangleShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1RectangleShape.html
.. _SubTitleShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1SubtitleShape.html
.. _TextRange: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextRange.html
.. _TitleTextShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1TitleTextShape.html
.. _XAnimationNode: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1animations_1_1XAnimationNode.html
.. _XController: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XController.html
.. _XDrawPage: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPage.html
.. _XDrawView: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawView.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
.. _XText: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XText.html
.. _XTextRange: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1text_1_1XTextRange.html
