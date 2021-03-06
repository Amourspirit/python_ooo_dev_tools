.. _ch04:

******************************************
Chapter 4. Listening, and Other Techniques
******************************************

.. topic:: Window Listeners;

    Window Listeners;

This chapter concludes the general introduction to Office programming by looking at several techniques that will reappear periodically in later chapters: the use of window listeners.

1. Listening to a Window
========================

There are many Listeners in office `(near 140 or so)`.
Using |dsearch|_ and running ``loapi comp -a -m 150 -t interface -s Listener`` shows 139 listeners.


On |lo_api|_ you can go to XEventListener_ that display a tree diagram at the top of the page, and you can click on a subclass box to jump to its documentation.

One nice syntactic feature of listeners is that almost all their names end with “Listener”.
This makes them easy to find when searching through indices of class names, such as the `Class Index <https://api.libreoffice.org/docs/idl/ref/classes.html>`_
or using |dsearch|_.

The top-level document window can be monitored for changes using XTopWindowListener_, which responds to modifications of the window`s state,
such as when it is opened, closed, minimized, and made active.

The following class from the |exlisten|_ example illustrates how to use the listener. ``DocWindow`` inherits from :py:class:`~.x_top_window_adapter.XTopWindowAdapter`
which is a simple adapter of XTopWindowListener_ class.

.. tabs::

    .. code-tab:: python

        #!/usr/bin/env python
        # coding: utf-8
        from __future__ import annotations
        import time
        import sys
        from typing import TYPE_CHECKING

        from ooodev.utils.lo import Lo
        from ooodev.utils.gui import GUI
        from ooodev.office.write import Write
        from ooodev.listeners.x_top_window_adapter import XTopWindowAdapter

        from com.sun.star.awt import XExtendedToolkit
        from com.sun.star.awt import XWindow

        if TYPE_CHECKING:
            # only need types in design time and not at run time.
            from com.sun.star.lang import EventObject

        class DocWindow(XTopWindowAdapter):
            def __init__(self) -> None:
                super().__init__()
                self.closed = False
                loader = Lo.load_office(Lo.ConnectPipe())
                self.tk = Lo.create_instance_mcf(XExtendedToolkit, "com.sun.star.awt.Toolkit")
                if self.tk is not None:
                    self.tk.addTopWindowListener(self)

                self.doc = Write.create_doc(loader=loader)

                GUI.set_visible(True, self.doc)
                # triggers 2 opened and 2 activated events

            def windowOpened(self, event: EventObject) -> None:
                """is invoked when a window is activated."""
                print("WL: Opened")
                xwin = Lo.qi(XWindow, event.Source)
                GUI.print_rect(xwin.getPosSize())

            def windowActivated(self, event: EventObject) -> None:
                """is invoked when a window is activated."""
                print("WL: Activated")
                print(f"  Titile bar: {GUI.get_title_bar()}")

            def windowDeactivated(self, event: EventObject) -> None:
                """is invoked when a window is deactivated."""
                print("WL: Minimized")

            def windowMinimized(self, event: EventObject) -> None:
                """is invoked when a window is iconified."""
                print("WL:  De-activated")

            def windowNormalized(self, event: EventObject) -> None:
                """is invoked when a window is deiconified."""
                print("WL: Normalized")

            def windowClosing(self, event: EventObject) -> None:
                """
                is invoked when a window is in the process of being closed.

                The close operation can be overridden at this point.
                """
                print("WL: Closing")

            def windowClosed(self, event: EventObject) -> None:
                """is invoked when a window has been closed."""
                print("WL: Closed")
                self.closed = True

            def disposing(self, event: EventObject) -> None:
                """
                gets called when the broadcaster is about to be disposed.

                """
                print("WL: Disposing")


|exlisten|_ example is also demonstrates how to keep a python script alive while office is running.

.. tabs::

    .. code-tab:: python

        def main_loop() -> None:
            dw = DocWindow()

            # while Writer is open, keep running the script unless specifically ended by user
            while 1:
                if dw.closed is True: # wait for windowClosed event to be raised
                    print("\nExiting by document close.\n")
                    break
                time.sleep(0.1)

        if __name__ == "__main__":
            print("Press 'ctl+c' to exit script early.")
            try:
                main_loop()
            except KeyboardInterrupt:
                # ctrl+c exitst the script earily
                print("\nExiting by user request.\n", file=sys.stderr)
                sys.exit(0)

The class implements seven methods for XTopWindowListener_, and disposing() inherited from XEventListener_.

The DocWindow object is made the listener for the window by accessing the XExtendedToolkit_ interface,
which is part of the Toolkit_ service. ``Toolkit`` is utilized by Office to create windows, and ``XExtendedToolkit``
adds three kinds of listeners: XTopWindowListener_, XFocusListener_, and the XKeyHandler_ listener.

When an event arrives at a listener method, one of the more useful things to do is to transform it into an XWindow_ instance:

.. tabs::

    .. code-tab:: python

        xwin = Lo.qi(XWindow, event.Source)

It's then possible to access details about the frame, such as its size.

Events are fired when :py:meth:`.GUI.set_visible` is called in the ``DocWindow()`` constructor.
An opened event is issued, followed by an activated event, triggering calls to ``windowOpened()`` and ``windowActivated()``.
Rather confusingly, both these methods are called twice.

.. note::

    If :py:meth:`.Lo.close_doc` were to be called, a single de-activated event is fired, but two closed events are issued.
    Consequently, there's a single call to ``windowDeactivated()`` and two to ``windowClosed()``.
    Strangely, there's no window closing event trigger of ``windowClosing()``, and :py:meth:`.Lo.close` doesn't cause ``disposing()`` to fire.

4.2 Office Manipulation
=======================

Although XTopWindowListener_ can detect the minimization and re-activation of the document window, it cant't trigger events.
Listeners Listen for events but can not trigger events.

I the :py:class:`~.gui.GUI` class there are a few methods for basic window manipulation.
For instance to activate a window use :py:meth:`~.gui.GUI.activate`, for min and max there is :py:meth:`~.gui.GUI.minimize` and :py:meth:`~.gui.GUI.maximize`,
:py:meth:`get_pos_size` for size and postiion.

There are other python libraries that can handel mouse and keyboard enulation such as `PyAutoGUI <https://pypi.org/project/PyAutoGUI/>`_ and `keyboard <https://pypi.org/project/keyboard/>`_.
|app_name_short| will leave it up to developers to implement window manipulation for their own use.


4.3 Detecting Office Termination
================================



Work in Progress...

.. |exlisten| replace:: Office Window Listener
.. _exlisten: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_listen

.. _XEventListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XEventListener.html
.. _XTopWindowListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XTopWindowListener.html
.. _XExtendedToolkit: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XExtendedToolkit.html
.. _XFocusListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XFocusListener.html
.. _XKeyHandler: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XKeyHandler.html
.. _XWindow: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XWindow.html
.. _ToolKit: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1Toolkit.html

.. include:: ../../resources/odev/links.rst