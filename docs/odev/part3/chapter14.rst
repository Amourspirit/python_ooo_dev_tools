.. _ch14:

*********************
Chapter 14. Animation
*********************

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

.. topic:: Overview

    Circle Movement; Line Rotation; Animating an Image; The Gallery Module.

    Examples: |animate_bike|_, |draw_picture|_, and |gallery_info|_.

|draw_picture_py|_ contains a call to ``_anim_shapes()`` which shows how to animate a circle and a line.
There's a second animation example in |animate_bike_py|_ which translates and rotates a bicycle image.
The chapter ends with a brief outline of the com.sun.star.gallery module.

.. _ch14_animate_circle_ln:

14.1 Animating a Circle and a Line
==================================

``_anim_shapes()`` in |draw_picture_py|_ implements two animation loops that work in a similar manner.
Inside a loop, a shape is drawn, the function (and program) sleeps for a brief period,
then the shape's position, size, or properties are updated, and the loop repeats.

The first animation loop moves a circle across the page from left to right, reducing its radius at the same time.
The second loop rotates a line counter-clockwise while changing its length. The ``_anim_shapes()`` code:

.. tabs::

    .. code-tab:: python

        # from draw_picture.py
        def _anim_shapes(self, curr_slide: XDrawPage) -> None:
            # reduce circle size and move to the right
            xc = 40
            yc = 150
            radius = 40
            circle = None
            for _ in range(20):
                # move right
                if circle is not None:
                    curr_slide.remove(circle)
                circle = Draw.draw_circle(slide=curr_slide, x=xc, y=yc, radius=radius)

                Lo.delay(200)
                xc += 5
                radius *= 0.95

            # rotate line counter-clockwise, and change length
            x2 = 140
            y2 = 110
            line = None
            for _ in range(25):
                if line is not None:
                    curr_slide.remove(line)
                line = Draw.draw_line(slide=curr_slide, x1=40, y1=100, x2=x2, y2=y2)
                x2 -= 4
                y2 -= 4

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The shape (circle or line) is changed by removing the current version from the page and inserting a new updated instance.
This means that a lot of objects are created and removed in a short amount of time. The alternative approach,
which retains the shape and only update its properties, is used in the bicycle animation explained next.

.. _ch14_animate_img:

14.2 Animating an Image
=======================

The |animate_bike|_ example moves a bicycle image to the right and rotates it counter-clockwise.
:numref:`ch14fig_animated_bike_and_shapes` shows the page after the animation has finished.

..
    figure 1

.. cssclass:: screen_shot invert

    .. _ch14fig_animated_bike_and_shapes:
    .. figure:: https://user-images.githubusercontent.com/4193389/200142865-c02cbed5-f147-4664-af28-46e2a6010d51.png
        :alt: Animated Bicycle and Shapes.
        :figclass: align-center

        :Animated Bicycle and Shapes.

The animation is performed by ``_animate_bike()``:

.. tabs::

    .. code-tab:: python

        # from anim_bicycle.py
        def _animate_bike(self, slide: XDrawPage) -> None:
            shape = Draw.draw_image(slide=slide, fnm=self._fnm_bike, x=60, y=100, width=90, height=50)

            pt = Draw.get_position(shape)
            angle = Draw.get_rotation(shape)
            print(f"Start Angle: {int(angle)}")
            for i in range(19):
                Draw.set_position(shape=shape, x=pt.X + (i * 5), y=pt.Y)  # move right
                Draw.set_rotation(shape=shape, angle=angle + (i * 5))  # rotates ccw
                Lo.delay(200)

            print(f"Final Angle: {int(Draw.get_rotation(shape))}")
            Draw.print_matrix(Draw.get_transformation(shape))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The animation loop in ``_animate_bike()`` is similar to the ones in ``anim_shapes()``, using :py:meth:`.Lo.delay` to space out changes over time.
However, instead of creating a new shape on each iteration, a single GraphicObjectShape_ is created by :py:meth:`.Draw.draw_image` before the loop starts.
Inside the loop, that shape’s position and orientation are repeatedly updated by :py:meth:`.Draw.set_position` and :py:meth:`.Draw.set_rotation`.

.. _ch14_draw_img:

14.2.1 Drawing the Image
------------------------

There are several versions of :py:meth:`.Draw.draw_image` the main one is:

.. tabs::

    .. code-tab:: python

        # represents draw_image() overloads in Draw Class (simplified)
        @classmethod
        def draw_image(cls, slide: XDrawPage, fnm: PathOrStr) -> XShape:
            slide_size = cls.get_slide_size(slide)
            im_size = ImagesLo.get_size_100mm(fnm)
            im_width = round(im_size.Width / 100)  # in mm units
            im_height = round(im_size.Height / 100)
            x = round((slide_size.Width - im_width) / 2)
            y = round((slide_size.Height - im_height) / 2)
            return cls.draw_image(slide=slide, fnm=fnm, x=x, y=y, width=im_width, height=im_height)

        @classmethod
        def draw_image(
            cls,
            slide: XDrawPage,
            fnm: PathOrStr,
            x: int,
            y: int,
            width: int,
            height: int
        ) -> XShape:

            # units in mm's
            Lo.print(f'Adding the picture "{fnm}"')
            im_shape = cls.add_shape(
                slide=slide,
                shape_type=DrawingShapeKind.GRAPHIC_OBJECT_SHAPE,
                x=x,
                y=y,
                width=width,
                height=height
            )
            cls.set_image(im_shape, fnm)
            cls.set_line_style(shape=im_shape, style=LineStyle.NONE)
            return im_shape

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`draw_image`

:py:meth:`~.Draw.draw_image` uses the supplied (x, y) position, width, and height to create an empty GraphicObjectShape_.
An image is added by ``setImage()``, which loads a bitmap from a file, and assigns it to the shape's ``GraphicURL`` property.
By using a bitmap, the image is embedded in the document.

Alternatively, a URL could be assigned to ``GraphicURL``, causing the document's image to be a link back to its original file.

That version is coded using:

.. tabs::

    .. code-tab:: python

        Props.set(GraphicURL=FileIO.fnm_to_url(im_fnm))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

A second version of :py:meth:`.Draw.draw_image` doesn't require width and height arguments – they're obtained from the image’s dimensions:

.. tabs::

    .. code-tab:: python

        # represents draw_image() overload in Draw Class (simplified)
        @classmethod
        def draw_image(cls, slide: XDrawPage, fnm: PathOrStr, x: int, y: int) -> XShape:
            im_size = ImagesLo.get_size_100mm(fnm)
            return cls.draw_image(
                slide=slide,
                fnm=fnm,
                x=x,
                y=y,
                width=round(im_size.Width / 100),
                height=round(im_size.Height / 100)
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`draw_image`

The image's size is returned in ``1/100 mm`` units by :py:meth:`.ImagesLo.get_size_100mm`.
It loads the image as an XGraphic_ object so that its ``Size100thMM`` property can be examined:

.. tabs::

    .. code-tab:: python

        # in the ImagesLo class
        @classmethod
        def get_size_100mm(cls, im_fnm: PathOrStr) -> Size:
            graphic = cls.load_graphic_file(im_fnm)
            return mProps.Props.get(graphic, "Size100thMM")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This approach isn't very efficient since it means that the image is being loaded twice,
once as an XGraphic_ object by :py:meth:`~.ImagesLo.get_size_100mm`, and also as a bitmap by ``setImage()``.

.. _ch14_update_bike:

14.2.2 Updating the Bike's Position and Orientation
---------------------------------------------------

The ``_animate_bike()`` animation uses Draw methods for getting and setting the shap's position and orientation:

.. tabs::

    .. code-tab:: python

        # in the Draw Class (simplified)
        @staticmethod
        def get_position(shape: XShape) -> Point:
            pt = shape.getPosition()
            # convert to mm
            return Point(round(pt.X / 100), round(pt.Y / 100))

        # one of several overloads
        @staticmethod
        def set_position(shape: XShape, x: int, y: int) -> None:
            shape.set_position(Point(x * 100, y * 100))
        
        @staticmethod
        def get_rotation(shape: XShape) -> Angle:
            r_angle = int(mProps.Props.get(shape, "RotateAngle"))
            return Angle(round(r_angle / 100))

        @staticmethod
        def set_rotation(shape: XShape, angle: Angle) -> None:
            mProps.Props.set(shape, RotateAngle=angle.Value * 100)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: src-link

        -  :odev_src_draw_meth:`get_position`
        -  :odev_src_draw_meth:`set_position`
        -  :odev_src_draw_meth:`get_rotation`
        -  :odev_src_draw_meth:`set_rotation`

The position is accessed and changed using the XShape_ methods :py:meth:`~.Draw.get_position` and :py:meth:`~.Draw.set_position`,
with the only complication being the changes of millimeters into ``1/100 mm`` units, and vice versa.

Rotation is handled by getting and setting the shape's ``RotateAngle`` property, which is inherited from the RotationDescriptor_ class.
The angle is expressed in ``1/100`` of a degree units (:abbreviation:`e.g.` 4500 rather than 45 degrees), and a positive rotation is counter-clockwise.

One issue is that RotationDescriptor_ is deprecated; the modern programmer is encouraged to rotate a shape using the matrix associated with the ``Transformation`` property.

The Draw class has are two support functions for ``Transformation``: one extracts the matrix from a shape, and the other prints it:

.. tabs::

    .. code-tab:: python

        # in the Draw Class (simplified)
        @staticmethod
        def get_transformation(shape: XShape) -> HomogenMatrix3:
            return mProps.Props.get(shape, "Transformation")

        @staticmethod
        def print_matrix(mat: HomogenMatrix3) -> None:
            print("Transformation Matrix:")
            print(f"\t{mat.Line1.Column1:10.2f}\t{mat.Line1.Column2:10.2f}\t{mat.Line1.Column3:10.2f}")
            print(f"\t{mat.Line2.Column1:10.2f}\t{mat.Line2.Column2:10.2f}\t{mat.Line2.Column3:10.2f}")
            print(f"\t{mat.Line3.Column1:10.2f}\t{mat.Line3.Column2:10.2f}\t{mat.Line3.Column3:10.2f}")

            rad_angle = math.atan2(mat.Line2.Column1, mat.Line1.Column1)
            #       sin(t), cos(t)
            curr_angle = round(math.degrees(rad_angle))
            print(f"  Current angle: {curr_angle}")
            print()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

These methods are called at the end of ``_animate_bike()``:

.. tabs::

    .. code-tab:: python

        # from anim_bicycle.py _animate_bike()
        Draw.print_matrix(Draw.get_transformation(shape))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The output is:

::

    Transformation Matrix:
              0.00         5001.00        15383.00
          -9001.00            0.00        10235.00
              0.00            0.00            1.00
      Current angle: -90

These numbers suggests that the transformation was a clockwise rotation, but the calls to :py:meth:`.Draw.set_rotation` in the earlier animation loop made the bicycle turn counter-clockwise.
This discrepancy pointed to stay with the deprecated approach for shape rotation.

.. _ch14_alt_gallery_access:

14.3 Another Way to Access the Gallery
======================================

There's an alternative way to obtain gallery images based around themes and items, implemented by the ``com.sun.star.gallery`` module.
Sub-directories of ``gallery/`` are themes, and the files in those directories are items.

The three interfaces in the module are: XGalleryThemeProvider_, XGalleryTheme_, and XGalleryItem_.
XGalleryThemeProvider_ represents the ``gallery/`` directory as a sequence of named XGalleryTheme_ objects, as shown in :numref:`ch14fig_gallery_theme_provider_service`.

..
    figure 2

.. cssclass:: diagram invert

    .. _ch14fig_gallery_theme_provider_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/200184070-b13d262f-829b-4562-b41f-c9d683e35b72.png
        :alt: The Gallery Theme Provider Service
        :figclass: align-center

        :The GalleryThemeProvider_ Service.

A XGalleryTheme_ represents the file contents of a sub-directory as a container of indexed XGalleryItem_ objects, which is depicted in :numref:`ch14fig_gallery_theme_service`.

..
    figure 3

.. cssclass:: diagram invert

    .. _ch14fig_gallery_theme_service:
    .. figure:: https://user-images.githubusercontent.com/4193389/200184471-fa856e68-ea7f-4395-b0c4-e1c23d271ae5.png
        :alt: The Gallery Theme Service
        :figclass: align-center

        :The GalleryTheme_ Service.

Each XGalleryItem_ represents a file, which may be a graphic or some other resource, such as an audio file.
The details about each item (file) are stored as properties which are defined in the GalleryItem_ service.

The :py:class:`~.utils.gallery.Gallery` class helps access the gallery in this way, and |gallery_info_py|_ contains some examples of its use:

.. tabs::

    .. code-tab:: python

        # from gallery_info.py
        from __future__ import annotations
        import uno
        from ooodev.utils.lo import Lo
        from ooodev.utils.gallery import Gallery, GalleryKind, SearchByKind, SearchMatchKind


        class GalleryInfo:
            def main(self) -> None:
                with Lo.Loader(Lo.ConnectPipe(headless=True)):
                    # list all the gallery themes (i.e. the sub-directories below gallery/)
                    Gallery.report_galleries()
                    print()

                    # list all the items for the Sounds theme
                    Gallery.report_gallery_items(GalleryKind.SOUNDS)
                    print()

                    # find an item that has "applause" as part of its name
                    # in the Sounds theme
                    itm = Gallery.find_gallery_obj(
                        gallery_name=GalleryKind.SOUNDS,
                        name="applause",
                        search_match=SearchMatchKind.PARTIAL_IGNORE_CASE,
                        search_kind=SearchByKind.FILE_NAME,
                    )
                    print()
                    # print out the item's properties
                    Gallery.report_gallery_item(itm)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Gallery.report_galleries` gives details about ``9`` themes, :py:meth:`.Gallery.report_gallery_items` prints the names of the ``35`` items (files) in the Sounds theme.

:py:meth:`.Gallery.find_gallery_obj` searches that theme for an item name containing ``applause``, and :py:meth:`.Gallery.report_gallery_item` reports its details:

::

    Searching gallery "Sounds" for "applause"
      Search is ignoring case
      Searching for a partial match

    Found matching item: applause.wav

    Gallery item information:
      URL: "file:///C:/Program%20Files/LibreOffice/share/gallery/sounds/applause.wav"
      Fnm: "applause.wav"
      Path: C:\Program Files\LibreOffice\share\gallery\sounds\applause.wav
      Title: ""
      Type: media

.. |animate_bike| replace:: Animate Bike
.. _animate_bike: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_animate_bike

.. |animate_bike_py| replace:: anim_bicycle.py
.. _animate_bike_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/draw/odev_animate_bike/anim_bicycle.py

.. |draw_picture| replace:: Draw Picture
.. _draw_picture: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_draw_picture

.. |draw_picture_py| replace:: draw_picture.py
.. _draw_picture_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/draw/odev_draw_picture/draw_picture.py

.. |gallery_info| replace:: Gallery Info
.. _gallery_info: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_gallery_info

.. |gallery_info_py| replace:: gallery_info.py
.. _gallery_info_py: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_gallery_info/gallery_info.py

.. _GalleryItem: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1gallery_1_1GalleryItem.html
.. _GalleryTheme: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1gallery_1_1GalleryTheme.html
.. _GalleryThemeProvider: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1gallery_1_1GalleryThemeProvider.html
.. _GraphicObjectShape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicObjectShape.html
.. _RotationDescriptor: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1RotationDescriptor.html
.. _XGalleryItem: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1gallery_1_1XGalleryItem.html
.. _XGalleryTheme: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1gallery_1_1XGalleryTheme.html
.. _XGalleryThemeProvider: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1gallery_1_1XGalleryThemeProvider.html
.. _XGraphic: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1graphic_1_1XGraphic.html
.. _XShape: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XShape.html
