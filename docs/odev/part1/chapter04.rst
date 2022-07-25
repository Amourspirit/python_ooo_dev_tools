.. _ch04:

******************************************
Chapter 4. Listening, and Other Techniques
******************************************

.. topic:: Window Listeners;

    Window Listeners;

This chapter concludes the general introduction to Office programming by looking at several techniques that will reappear periodically in later chapters: the use of window listeners.

.. _ch04_listen_win:

4.1 Listening to a Window
=========================

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

.. _ch04_office_manipulate:

4.2 Office Manipulation
=======================

Although XTopWindowListener_ can detect the minimization and re-activation of the document window, it can't trigger events.
Listeners Listen for events but can not trigger events.

I the :py:class:`~.gui.GUI` class there are a few methods for basic window manipulation.
For instance to activate a window use :py:meth:`~.gui.GUI.activate`, for min and max there is :py:meth:`~.gui.GUI.minimize` and :py:meth:`~.gui.GUI.maximize`,
:py:meth:`get_pos_size` for size and position.

There are other python libraries that can handle mouse and keyboard emulation such as `PyAutoGUI <https://pypi.org/project/PyAutoGUI/>`_ and `keyboard <https://pypi.org/project/keyboard/>`_.
|odev| will leave it up to developers to implement window manipulation for their own use.

.. _ch04_detect_end:

4.3 Detecting Office Termination
================================

Office termination is most easily observed by attaching a listener to the Desktop object,
as seen in |exmonitor|_ example.

.. tabs::

    .. code-tab:: python

        class DocMonitor:

            def __init__(self) -> None:
                super().__init__()
                self.closed = False
                loader = Lo.load_office(Lo.ConnectPipe())
                xdesktop = Lo.XSCRIPTCONTEXT.getDesktop()
                
                # create a new instance of adapter. Note that adapter methods all pass.
                term_adapter = XTerminateAdapter()
                
                # reassign the method we want to use from XTerminateAdapter instance in a pythonic way.
                term_adapter.notifyTermination = types.MethodType(self.notify_termination, term_adapter)
                term_adapter.queryTermination = types.MethodType(self.query_termination, term_adapter)
                term_adapter.disposing = types.MethodType(self.disposing, term_adapter)
                xdesktop.addTerminateListener(term_adapter)
                
                self.doc = Calc.create_doc(loader=loader)

                GUI.set_visible(True, self.doc)


            def notify_termination(self, src: XTerminateAdapter, event: EventObject) -> None:
                """
                is called when the master environment is finally terminated.
                """
                print("TL: Finished Closing")
                self.closed = True
            
            def query_termination(self, src: XTerminateAdapter, event: EventObject) -> None:
                """
                is called when the master environment (e.g., desktop) is about to terminate.
                """
                print("TL: Starting Closing")
                
                
            def disposing(self, src: XTerminateAdapter, event: EventObject) -> None:
                """
                gets called when the broadcaster is about to be disposed.
                """
                print("TL: Disposing")

An XTerminateListener is attached to the XDesktop instance. The program's output is:

.. code-block:: text

    PS D:\Users\user\Python\python-ooouno-ex> python -m main auto --process "ex/auto/general/odev_monitor/start.py True"
    Press 'ctl+c' to exit script early.
    Loading Office...
    Creating Office document scalc
    Closing Office
    TL: Starting Closing
    TL: Finished Closing
    Office terminated

    Exiting by document close.

XTerminateListener_’s ``queryTermination()`` and ``notifyTermination()`` are called at the start and end of the Office closing sequence.
As in the |exlisten|_ example, ``disposing()`` is never triggered.

.. _ch04_bridge_stop:

4.4 Bridge Shutdown Detection
=============================

There's another way to detect Office closure: by listening for the shutdown of the UNO bridge between the Python and Office processes.
This can be useful if Office crashes independently of your Python code.
This approach works for both socket and pipe connections using python.

The modified parts of |exmonitor|_ are:

.. tabs::

    .. code-tab:: python

        # new imports
        from ooodev.listeners.x_event_adapter import XEventAdapter

        # DocMonitor constructor changes
        self.bridge_disposed = False

        bridge_listen = XEventAdapter()
        bridge_listen.disposing = types.MethodType(self.disposing_bridge, bridge_listen)
        Lo.bridge.addEventListener(bridge_listen)

        # DocMonitor new method
        def disposing_bridge(self, src:XEventAdapter, event:EventObject) -> None:
            print("BR: Office bridge has gone!!")
            self.bridge_disposed = True

        # main_loop method adds check for dw.bridge_disposed
        def main_loop() -> None:
            dw = DocMonitor()

            # check and see if user passed in a auto terminate option
            if len(sys.argv) > 1:
                if str(sys.argv[1]).casefold() in ("t", "true", "y", "yes"):
                    Lo.delay(5000)
                    Lo.close_office()

            while 1:
                if dw.closed is True:  # wait for windowClosed event to be raised
                    print("\nExiting by document close.\n")
                    break
                if dw.bridge_disposed is True:
                    print("\nExiting due to office bridge is gone\n")
                    raise SystemExit(1)
                time.sleep(0.1)

Since the disappearance of the Office bridge is a fatal event, in ``main_loop()``  ``raise SystemExit(1)`` is called to kill python.

.. note::

    Raising ``SystemExit(1)`` inside of ``disposing_bridge()`` does not exit script and that is why it is delegated to ``main_loop()``

The output of the revised |exmonitor|_ is:

.. code-block:: text

    PS D:\Users\user\Python\python-ooouno-ex> python -m main auto --process "ex/auto/general/odev_monitor/start.py True"
    Press 'ctl+c' to exit script early.
    Loading Office...
    Creating Office document scalc
    Closing Office
    TL: Starting Closing
    TL: Finished Closing
    Office terminated
    BR: Office bridge has gone!!

    Exiting by document close.

This output shows that bridge closure follows the call to :py:meth:`.Lo.close_office`, as you'd expect.
However, if I make Office crash while DocMonitor is running, then the output becomes:

.. code-block:: text

    PS D:\Users\user\Python\python-ooouno-ex> python -m main auto --process "ex/auto/general/odev_monitor/start.py"
    Press 'ctl+c' to exit script early.
    Loading Office...
    Creating Office document scalc
    Office bridge has gone!!
    BR: Office bridge has gone!!

    Exiting due to office bridge is gone

Office was killed while the python program was still running, so it never reached its :py:meth:`.Lo.close_office` call which triggers the XTerminateListener_ methods.
However, the XEventListener_ attached to the bridge did fire.
(If you're wondering, office was killed Office by running ``loproc -k``, which stopped the soffice process. See: |dsearch|_)

.. _ch04_dispatching:

4.5 Dispatching
===============

This book is about the Python Office API, which manipulates UNO data structures such as services, interfaces, and components.
There's an alternative programming style, based on the dispatching of messages to Office.
These messages are mostly related to menu items, so, for example, the messages ``.uno:Copy``, ``.uno:Cut``, ``.uno:Paste``, and ``.uno:Undo`` duplicate commands in the **Edit** menu.
The use of messages tends to be most common when writing macros (scripts) in Basic,
because Office's built-in Macro recorder automatically converts a user's interaction with menus into dispatches.

One drawback of dispatching is that it isn't a complete programming solution.
For instance, copying requires the selection of text, which has to be implemented with the Office API.

LibreOffice has a comprehensive webpage listing all the dispatch commands `Development/DispatchCommands <https://wiki.documentfoundation.org/Development/DispatchCommands>`_.

Another resource is chapter 10 of `OpenOffice.org Macros Explained <https://pitonyak.org/book/>`_ by Andrew Pitonyak.

Creating a dispatcher in Python is a little complicated since XDispatchProvider_ and XDispatchHelper_ instances are needed.
XDispatchProvider_ is obtained from the frame (i.e. window) where the message will be delivered, which is almost always the Desktop's frame
(i.e. Office application's window). XDispatchHelper_ sends the message via its ``executeDispatch()`` method.
It's also possible to examine the result status in an DispatchResultEvent_ object, but that seems a bit flakey – it reports failure when the dispatch works,
and raises an exception when the dispatch really fails.

The code is wrapped up in :py:meth:`.Lo.dispatch_cmd`, which is called twice in the |ex_dispatch|_:

.. collapse:: Example
    :open:

    .. tabs::

        .. code-tab:: python

            #!/usr/bin/env python
            # coding: utf-8
            # region Imports
            from __future__ import annotations
            import argparse
            import sys

            from ooodev.utils.lo import Lo
            from ooodev.utils.gui import GUI

            # endregion Imports

            # region args
            def args_add(parser: argparse.ArgumentParser) -> None:
                parser.add_argument(
                    "-d",
                    "--doc",
                    help="Path to document",
                    action="store",
                    dest="fnm_doc",
                    required=True,
                )


            # endregion args

            # region Main
            def main() -> int:
                # create parser to read terminal input
                parser = argparse.ArgumentParser(description="main")

                # add args to parser
                args_add(parser=parser)

                if len(sys.argv) <= 1:
                    parser.print_help(sys.stderr)
                    return 1

                # read the current command line args
                args = parser.parse_args()

                fnm = args.fnm_doc
                loader = Lo.load_office(Lo.ConnectPipe())
                try:
                    doc = Lo.open_doc(fnm=fnm, loader=loader)
                    # breakpoint()
                except Exception:
                    print(f"Could not open '{fnm}'")
                    Lo.close_office()
                    raise SystemExit(1)

                GUI.set_visible(is_visible=True, odoc=doc)
                Lo.delay(3000)  # delay 3 seconds

                # put doc into readonly mode
                Lo.dispatch_cmd("ReadOnlyDoc")
                Lo.delay(1000)

                # opens get involved webpage of LibreOffice in local browser
                Lo.dispatch_cmd("GetInvolved")

                return 0


            if __name__ == "__main__":
                raise SystemExit(main())

            # endregion main


It is also possible in |odev| to hook events. These events are specific to |odev| and not
part of LibreOffice.

Here is the updated example that hooks DISPATCHING and DISPATCHED events.
In the example the ``About`` dispatch is canceled.

See :ref:`events_lo_events_Events`.

.. collapse:: Extended Example
    :open:

    .. tabs::

        .. code-tab:: python

            #!/usr/bin/env python
            # coding: utf-8
            # region Imports
            from __future__ import annotations
            import argparse
            import sys
            from typing import Any

            from ooodev.utils.lo import Lo
            from ooodev.utils.gui import GUI
            from ooodev.events.lo_events import Events
            from ooodev.events.lo_named_event import LoNamedEvent
            from ooodev.events.args.dispatch_args import DispatchArgs
            from ooodev.events.args.dispatch_cancel_args import DispatchCancelArgs

            # endregion Imports

            # region args
            def args_add(parser: argparse.ArgumentParser) -> None:
                parser.add_argument(
                    "-d",
                    "--doc",
                    help="Path to document",
                    action="store",
                    dest="fnm_doc",
                    required=True,
                )


            # endregion args

            # region dispatch events
            def on_dispatching(source: Any, event: DispatchCancelArgs) -> None:
                if event.cmd == "About":
                    print("About dispatch canceled")
                    event.cancel = True
                    return
                print(f"Dispatching: {event.cmd}")


            def on_dispatched(source: Any, event: DispatchArgs) -> None:
                print(f"Dispatched: {event.cmd}")


            # endregion dispatch events

            # region Main
            def main() -> int:
                # create parser to read terminal input
                parser = argparse.ArgumentParser(description="main")

                # add args to parser
                args_add(parser=parser)

                if len(sys.argv) <= 1:
                    parser.print_help(sys.stderr)
                    return 1

                # read the current command line args
                args = parser.parse_args()

                fnm = args.fnm_doc
                loader = Lo.load_office(Lo.ConnectPipe())
                try:
                    doc = Lo.open_doc(fnm=fnm, loader=loader)
                    # breakpoint()
                except Exception:
                    print(f"Could not open '{fnm}'")
                    Lo.close_office()
                    raise SystemExit(1)

                # create an instance of events to hook into ooodev events
                events = Events()
                events.on(LoNamedEvent.DISPATCHING, on_dispatching)
                events.on(LoNamedEvent.DISPATCHED, on_dispatched)

                GUI.set_visible(is_visible=True, odoc=doc)
                Lo.delay(3000)  # delay 3 seconds

                # put doc into readonly mode
                Lo.dispatch_cmd("ReadOnlyDoc")
                Lo.delay(1000)

                # opens get involved webpage of LibreOffice in local browser
                Lo.dispatch_cmd("GetInvolved")
                Lo.dispatch_cmd("About")

                return 0


            if __name__ == "__main__":
                raise SystemExit(main())

            # endregion main

Here is the output from extended example.

.. code-block:: text

    PS D:\Users\user\Python\python-ooouno-ex> python -m main auto --process 'ex\auto\writer\odev_dispatch\start.py -d "resources\odt\story.odt"'
    Loading Office...
    Opening D:\Users\user\Python\python-ooouno-ex\resources\odt\story.odt
    Dispatching: ReadOnlyDoc
    Dispatched: ReadOnlyDoc
    Dispatching: GetInvolved
    Dispatched: GetInvolved
    About dispatch canceled

.. |ex_dispatch| replace:: Dispatch Commands Example
.. _ex_dispatch: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_dispatch
.. |exlisten| replace:: Office Window Listener
.. _exlisten: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_listen

.. |exmonitor| replace:: Office Window Monitor
.. _exmonitor: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_monitor

.. _XDispatchProvider: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XDispatchProvider.html
.. _XDispatchHelper: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XDispatchHelper.html
.. _DispatchResultEvent: https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1frame_1_1DispatchResultEvent.html
.. _XEventListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XEventListener.html
.. _XTopWindowListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XTopWindowListener.html
.. _XTerminateListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XTerminateListener.html
.. _XExtendedToolkit: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XExtendedToolkit.html
.. _XFocusListener: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XFocusListener.html
.. _XKeyHandler: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XKeyHandler.html
.. _XWindow: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XWindow.html
.. _ToolKit: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1Toolkit.html

.. include:: ../../resources/odev/links.rst