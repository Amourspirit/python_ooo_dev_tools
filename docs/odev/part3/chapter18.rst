.. _ch18:

***********************
Chapter 18. Slide Shows
***********************

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

.. topic:: Overview

    Starting a Slide Show; Play and End a Slide Show Automatically; Play a Slide Show Repeatedly; Play a Custom Slide Show

     Examples: |auto_show|_, |basic_show|_, and |c_show|_. 

This chapter focuses on ways to programmatically control slide shows.
If you're unfamiliar with what Impress offers in this regard, then chapter 9 of the Impress user guide gives a good overview.

Creating and controlling slide shows employs properties in the Presentation service, and methods in the XSlideShowController_ interface (see :numref:`ch18fig_ss_pres_services`).

..
    figure 1

.. cssclass:: diagram invert

    .. _ch18fig_ss_pres_services:
    .. figure:: https://user-images.githubusercontent.com/4193389/202236249-f0fc428f-66d3-4f6d-930e-c93c0d4b6cab.png
        :alt: The Slide Show Presentation Services
        :figclass: align-center

        :The Slide Show Presentation Services.

Two elements of slide shows not shown in :numref:`ch18fig_ss_pres_services` are slide transition effects (:abbreviation:`i.e.` have the next slide fade into view, replacing the current one),
and shape animation effects (:abbreviation:`i.e.` have some text whoosh in from the bottom of the screen). These effects are mostly controlled by setting properties – transition properties
are in the |p_drawpage|_ service, animations properties in |p_shape|_.

.. _ch18_starting_slide_show:

18.1 Starting a Slide Show
==========================

The |basic_show_py|_ example shows how a program can start a slide show, and then let the user progress through the presentation by clicking on a slide,
pressing the space bar, or using the arrow keys.

While the slide show is running, |basic_show_py|_ suspends, but wakes up when the user exits the show.
This can occur when he presses the ESC key, or clicks on the slide show's "click to exit" screen.
|basic_show_py|_ then closes the document and shuts down Office.

The |basic_show_py|_ module:

.. tabs::

    .. code-tab:: python

        # basic_show.py module.
        from __future__ import annotations
        import uno
        from ooodev.office.draw import Draw
        from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
        from ooodev.utils.file_io import FileIO
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.utils.props import Props
        from ooodev.utils.type_var import PathOrStr


        class BasicShow:
            def __init__(self, fnm: PathOrStr) -> None:
                _ = FileIO.is_exist_file(fnm=fnm, raise_err=True)
                self._fnm = FileIO.get_absolute_path(fnm)

            def main(self) -> None:
                with Lo.Loader(Lo.ConnectPipe()) as loader:
                    doc = Lo.open_doc(fnm=self._fnm, loader=loader)
                    try:
                        # slideshow start() crashes if the doc is not visible
                        GUI.set_visible(is_visible=True, odoc=doc)

                        show = Draw.get_show(doc=doc)
                        Props.show_obj_props("Slide show", show)

                        Lo.delay(500)
                        Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)
                        # show.start() starts slideshow but not necessarily in 100% full screen
                        # show.start()

                        sc = Draw.get_show_controller(show)
                        Draw.wait_ended(sc)

                    finally:
                        Lo.close_doc(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The document is opened in the normal way and a slide show object created by calling :py:meth:`.Draw.get_show`, which is defined as:

.. tabs::

    .. code-tab:: python

        # in the Draw class
        @staticmethod
        def get_show(doc: XComponent) -> XPresentation2:
            try:
                ps = Lo.qi(XPresentationSupplier, doc, True)
                return Lo.qi(XPresentation2, ps.getPresentation(), True)
            except Exception as e:
                raise DrawError("Unable to get Presentation") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The call to :py:meth:`.Props.show_obj_props` in ``main()`` prints the properties associated with the slide show, most of which are defined in the
Presentation_ service (see :numref:`ch18fig_ss_pres_services`):

.. cssclass:: rst-collapse

    .. collapse:: Output:
        :open:

        ::

            Slide show Properties
              AllowAnimations: True
              CustomShow: 
              Display: 0
              FirstPage: 
              IsAlwaysOnTop: False
              IsAutomatic: False
              IsEndless: False
              IsFullScreen: True
              IsMouseVisible: False
              IsShowAll: True
              IsShowLogo: False
              IsTransitionOnClick: True
              Pause: 0
              StartWithNavigator: False
              UsePen: False

The default values for these properties are sufficient for most presentations.

The slide show can be started by calling ``XPresentation.show()``, However; this can start the presentation with the toolbars still showing.
For this reason starting with dispatch command (``Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)``) seemed the best option.
Although the call returns immediately, it may be a few 100 milliseconds before the presentation appears on screen.
If you have more than one monitor, one of them will be allocated a "Presenter Console" window.

This short period while the slide show initializes can cause a problem if the XSlideShowController_ instance is requested too quickly – ``None`` will be returned
if the slide show hasn't finished being created. :py:meth:`.Draw.get_show_controller` handles this issue by waiting:

.. tabs::

    .. code-tab:: python

        # in the Draw class
        @staticmethod
        def get_show_controller(show: XPresentation2) -> XSlideShowController:
            try:
                sc = show.getController()
                # may return None if executed too quickly after start of show
                if sc is not None:
                    return sc
                timeout = 5.0  # wait time in seconds
                try_sleep = 0.5
                end_time = time.time() + timeout
                while end_time > time.time():
                    time.sleep(try_sleep)  # give slide show time to start
                    sc = show.getController()
                    if sc is not None:
                        break
            except Exception as e:
                raise DrawError("Error getting slide show controller") from e
            if sc is None:
                raise DrawError(f"Could obtain slide show controller after {timeout:.1f} seconds")
            return sc

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Draw.get_show_controller` tries to obtain the controller for ``5`` seconds before giving up and raising :py:class:`~.ex.exceptions.DrawError`.

The XSlideShowController_ interface gives the programmer much greater control over the slide show,
including the ability to change the slide being displayed, and monitor and control the slide show state.
Two topics that are not covered here are how XSlideShowController_ can assign listeners to the slide show (of type XSlideShowListener_), and how to utilize the XSlideShow_ interface.

Back in |basic_show_py|_, the ``main()`` function suspends by calling :py:meth:`.Draw.wait_ended`;
the idea is that the program will sleep until the human presenter ends the slide show.
:py:meth:`~.Draw.wait_ended` is implemented using XSlideShowController_:

.. tabs::

    .. code-tab:: python

        # in the Draw Class
        @staticmethod
        def wait_ended(sc: XSlideShowController) -> None:
            while True:
                curr_index = sc.getCurrentSlideIndex()
                if curr_index == -1:
                    break
                Lo.delay(500)

            Lo.print("End of presentation detected")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

``XSlideShowController.getCurrentSlideIndex()`` normally returns a slide index (:abbreviation:`i.e.` ``0`` or greater), but when the slide show has finished it returns ``-1``.
:py:meth:`~.Draw.wait_ended` keeps polling for this value, sleeping for half a second between each test.

.. _ch18_play_and_end_slideshow:

18.2 Play and End a Slide Show Automatically
============================================

The |auto_show_py|_ example removes the need for a presenter to click on a slide to progress to the next one, and terminates the show itself after the last slide had been displayed:

.. tabs::

    .. code-tab:: python

        # in auto_show.py
        def main(self) -> None:
            loader = Lo.load_office(Lo.ConnectPipe())

            try:
                doc = Lo.open_doc(self._fnm, loader)

                # slideshow start() crashes if the doc is not visible
                GUI.set_visible(is_visible=True, odoc=doc)

                # set up a fast automatic change between all the slides
                slides = Draw.get_slides_list(doc)
                for slide in slides:
                    Draw.set_transition(
                        slide=slide,
                        fade_effect=self._fade_effect,
                        speed=AnimationSpeed.FAST,
                        change=DrawingSlideShowKind.AUTO_CHANGE,
                        duration=self._duration,
                    )

                show = Draw.get_show(doc)
                Props.show_obj_props("Slide Show", show)
                self._set_show_prop(show)
                # Props.set(show, IsEndless=True, Pause=0)

                Lo.delay(500)
                Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)
                # show.start() starts slideshow but not necessarily in 100% full screen
                # show.start()

                sc = Draw.get_show_controller(show)
                Draw.wait_last(sc=sc, delay=self._end_delay)
                Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION_END)
                Lo.delay(500)

                msg_result = MsgBox.msgbox(
                    "Do you wish to close document?",
                    "All done",
                    boxtype=MessageBoxType.QUERYBOX,
                    buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
                )
                if msg_result == MessageBoxResultsEnum.YES:
                    print("Ending the slide show")
                    Lo.close_doc(doc=doc, deliver_ownership=True)
                    Lo.close_office()
                else:
                    print("Keeping document open")
            except Exception:
                Lo.close_office()
                raise

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _ch18_automatic_transition:

Automatic Slide Transitioning
-----------------------------

The automated transition between slides is configured by calling :py:meth:`.Draw.set_transition` on every slide in the deck:

.. tabs::

    .. code-tab:: python

        # in AutoShow.main() of auto_show.py
        Draw.set_transition(
            slide=slide,
            fade_effect=self._fade_effect,
            speed=AnimationSpeed.FAST,
            change=DrawingSlideShowKind.AUTO_CHANGE,
            duration=self._duration,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Draw.set_transition` combines the setting of four slide properties: ``Effect``, ``Speed``, ``Change``, and ``Duration``:

.. tabs::

    .. code-tab:: python

        # in Draw class
        @staticmethod
        def set_transition(
            slide: XDrawPage,
            fade_effect: FadeEffect,
            speed: AnimationSpeed,
            change: DrawingSlideShowKind,
            duration: int,
        ) -> None:
            try:
                ps = Lo.qi(XPropertySet, slide, True)
                ps.setPropertyValue("Effect", fade_effect)
                ps.setPropertyValue("Speed", speed)
                ps.setPropertyValue("Change", int(change))
                # if change is SlideShowKind.AUTO_CHANGE
                # then Duration is used
                ps.setPropertyValue("Duration", abs(duration))  # in seconds
            except Exception as e:
                raise DrawPageError("Could not set slide transition") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Slide transition properties (such as ``Effect``, ``Speed``, ``Change``, and ``Duration``) are defined in the |p_drawpage|_ service.
However, the possible values for ``Effect`` are stored in an enumeration listed at the end of the |p_module|_ module :numref:`ch18fig_fade_effect_enum` shows the FadeEffect_ enum.

..
    figure 2

.. cssclass:: screen_shot invert

    .. _ch18fig_fade_effect_enum:
    .. figure:: https://user-images.githubusercontent.com/4193389/202278483-62bbd186-a6dd-4413-81c3-e17dccce4b25.png
        :alt: The FadeEffect Enum
        :figclass: align-center

        :The FadeEffect_ Enum.

The ``Speed`` property of AnimationSpeed_ is used to set the speed of a slide transition.
There are three possible settings: ``SLOW``, ``MEDIUM``, and ``FAST``.

The ``Change`` property specifies how a transition is triggered.
The property can take one of three integer values, which aren't defined by LibreOffice as an enum so |odev| defines them as :py:class:`~.kind.drawing_slide_show_kind.DrawingSlideShowKind`.

The default behavior is represented by ``0`` (:py:attr:`.DrawingSlideShowKind.CLICK_ALL_CHANGE`) which requires the presenter to click on a slide to change it,
and a click is also need to trigger any shape animations on the page. A value of ``2`` (:py:attr:`.DrawingSlideShowKind.CLICK_PAGE_CHANGE`)
relieves the presenter from clicking to trigger shape animations, but he still needs to activate a slide transition manually.
|auto_show_py|_ a passes :py:attr:`.DrawingSlideShowKind.AUTO_CHANGE` to :py:meth:`.Draw.set_transition` which causes all the animations and transitions to execute automatically.

The ``Duration`` property is specified in seconds and refers to how long the current slide stays on display before the transition effect begins.
This is different from the ``Speed`` property which refers to how quickly a transition is performed.

.. _ch18_automatic_finish:

Finishing Automatically
-----------------------

The other aspect of this automated slide show is making it stop when the last slide has been displayed.
This is implemented by :py:meth:`.Draw.wait_last`:

.. tabs::

    .. code-tab:: python

        # in Draw class
        @staticmethod
        def wait_last(sc: XSlideShowController, delay: int) -> None:
            wait = int(delay)
            num_slides = sc.getSlideCount()
            while True:
                curr_index = sc.getCurrentSlideIndex()
                if curr_index == -1:
                    break
                if curr_index >= num_slides - 1:
                    break
                Lo.delay(500)

            if wait > 0:
                Lo.delay(wait)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`~.Draw.wait_last` keeps checking the current slide index and sleeps until the last slide in the deck is reached.
It then goes to sleep one last time, to give the final slide time to be seen by the user.

.. _ch18_play_repeat_show:

18.3 Play a Slide Show Repeatedly
=================================

Another common kind of automated slide show is one that plays the show repeatedly, only terminating when the presenter steps in and presses the ``ESC`` key.
This only requires a few lines to be changed in |auto_show_py|_, shown in below:

.. tabs::

    .. code-tab:: python

        # in auto_show.py
        # ...
        show = Draw.get_show(doc)
        Props.showObjProps("Slide show", show);
        Props.set(show, IsEndless=True, Pause=0);
        show.start()

        sc = Draw.get_show_controller(show)
        Draw.wait_ended(sc)
        print("Ending the slide show")
        sc.deactivate()
        show.end()
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The ``IsEndless`` property turns on slide show cycling, and ``Pause`` indicates how long the black "Click to exit" screen is displayed before the show restarts.

:py:meth:`.Draw.wait_ended` is the same as before – it makes |auto_show_py|_ suspend until the user clicks on the exit screen or presses the ``ESC`` key.

.. _ch18_play_custom_show:

18.4 Play a Custom Slide Show
-----------------------------

A custom slide show is a display sequence other than the usual one that starts with the first slide and moves linearly through to the last.
A named 'play list' of pages must be created, consisting of references to slides in the deck.
The list can point to the slides in any order, and may reference a slide more than once.

:py:meth:`.Draw.build_play_list` creates the named play list using three arguments: the slide document, an array of slide indices which represents the intended playing sequence, and a name for the list.
For example:

.. tabs::

    .. code-tab:: python

        play_list = Draw.build_play_list(doc, "ShortPlay", 5, 6, 7, 8)  # XNameContainer

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

This creates a play list called "ShortPlay" which will show the slides with indices ``5``, ``6``, ``7``, and ``8`` (note: the first slide has index ``0``).
:py:meth:`.Draw.build_play_list` is used in the |c_show_py|_ example:

.. tabs::

    .. code-tab:: python

        # custom_show.py module
        from __future__ import annotations
        import uno
        from ooodev.dialog.msgbox import (
            MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
        )
        from ooodev.office.draw import Draw
        from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
        from ooodev.utils.file_io import FileIO
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.utils.props import Props
        from ooodev.utils.type_var import PathOrStr


        class CustomShow:
            def __init__(self, fnm: PathOrStr, *slide_idx: int) -> None:
                FileIO.is_exist_file(fnm=fnm, raise_err=True)
                self._fnm = FileIO.get_absolute_path(fnm)
                for idx in slide_idx:
                    if idx < 0:
                        raise IndexError("Index cannot be negative")
                self._idxs = slide_idx

            def main(self) -> None:
                loader = Lo.load_office(Lo.ConnectPipe())

                try:
                    doc = Lo.open_doc(fnm=self._fnm, loader=loader)
                    # slideshow start() crashes if the doc is not visible
                    GUI.set_visible(is_visible=True, odoc=doc)

                    if len(self._idxs) > 0:
                        _ = Draw.build_play_list(doc, "ShortPlay", *self._idxs)
                        show = Draw.get_show(doc=doc)
                        Props.set(show, CustomShow="ShortPlay")
                        Props.show_obj_props("Slide show", show)
                        Lo.delay(500)
                        Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)
                        # show.start() starts slideshow but not necessarily in 100% full screen
                        # show.start()
                        sc = Draw.get_show_controller(show)
                        Draw.wait_ended(sc)

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
                    else:
                        MsgBox.msgbox(
                            "There were no slides indexes to create a slide show.",
                            "No Slide Indexes",
                            boxtype=MessageBoxType.WARNINGBOX,
                        )

                except Exception:
                    Lo.close_office()
                    raise

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The play list is installed by setting the ``CustomShow`` property in the slide show.
The rest of the code in |c_show_py|_ is similar to the |basic_show_py|_ example.

.. _ch18_play_list:

Creating a Play List Using Containers
-------------------------------------

The most confusing part of :py:meth:`.Draw.build_play_list` is its use of two containers to hold the play list:

.. tabs::

    .. code-tab:: python

        # part of the build_play_list in draw class
        # ...
        # get name container for the slide show
        play_list = cls.get_play_list(doc)

        # get factory from the container
        xfactory = Lo.qi(XSingleServiceFactory, play_list, True)

        # use factory to make an index container
        slides_con = Lo.qi(XIndexContainer, xfactory.createInstance(), True)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

An index container is created by ``XSingleServiceFactory.createInstance()``, which requires a factory instance.
This factory is most conveniently obtained from an existing container, namely the one for the slide show.
That's obtained by :py:meth:`.Draw.get_play_list`:

.. tabs::

    .. code-tab:: python

        # in the Draw class
        @staticmethod
        def get_play_list(doc: XComponent) -> XNameContainer:
            try:
                cp_supp = Lo.qi(XCustomPresentationSupplier, doc, True)
                return cp_supp.getCustomPresentations()
            except Exception as e:
                raise DrawError("Error getting play list") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

:py:meth:`.Draw.build_play_list` fills the index container with references to the slides, and then places it inside the name container:

.. tabs::

    .. code-tab:: python

        # in the Draw class
        @classmethod
        def build_play_list(cls, doc: XComponent, custom_name: str, *slide_idxs: int) -> XNameContainer:
            play_list = cls.get_play_list(doc)
            try:
                xfactory = Lo.qi(XSingleServiceFactory, play_list, True)
                slides_con = Lo.qi(XIndexContainer, xfactory.createInstance(), True)

                Lo.print("Building play list using:")
                j = 0
                for i in slide_idxs:
                    try:
                        slide = cls._get_slide_doc(doc, i)
                    except IndexError as ex:
                        Lo.print(f"  Error getting slide for playlist. Skipping index {i}")
                        Lo.print(f"    {ex}")
                        continue
                    slides_con.insertByIndex(j, slide)
                    j += 1
                    Lo.print(f"  Slide No. {i+1}, index: {i}")

                play_list.insertByName(custom_name, slides_con)
                Lo.print(f'Play list stored under the name: "{custom_name}"')
                return play_list
            except Exception as e:
                raise DrawError("Unable to build play list.") from e

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The for-loop employs the tuple of indices to get references to the slides via :py:meth:`.Draw.get_slide`.
Each reference is added to the index container.


.. |p_drawpage| replace:: com.sun.star.presentation.DrawPage
.. _p_drawpage: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1DrawPage.html

.. |p_shape| replace:: com.sun.star.presentation.Shape
.. _p_shape: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1Shape.html

.. |p_module| replace:: com.sun.star.presentation
.. _p_module: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1presentation.html


.. |basic_show| replace:: Basic Slide Show
.. _basic_show: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_basic_show

.. |basic_show_py| replace:: basic_show.py
.. _basic_show_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/impress/odev_basic_show/basic_show.py

.. |auto_show| replace:: Auto Slide Show
.. _auto_show: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_auto_show

.. |auto_show_py| replace:: auto_show.py
.. _auto_show_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/impress/odev_auto_show/auto_show.py

.. |c_show| replace:: Custom Slide Show
.. _c_show: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/impress/odev_custom_show

.. |c_show_py| replace:: custom_show.py
.. _c_show_py: https://github.com/Amourspirit/python-ooouno-ex/blob/main/ex/auto/impress/odev_custom_show/custom_show.py

.. _AnimationSpeed: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1presentation.html#a07b64dc4a366b20ad5052f974ffdbf62
.. _FadeEffect: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1presentation.html#a9db0b8c5e72e0ae290ff76da0dd53e3d
.. _Presentation: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1Presentation.html
.. _XSlideShow: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1presentation_1_1XSlideShow.html
.. _XSlideShowController: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1presentation_1_1XSlideShowController.html
.. _XSlideShowListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1presentation_1_1XSlideShowListener.html
