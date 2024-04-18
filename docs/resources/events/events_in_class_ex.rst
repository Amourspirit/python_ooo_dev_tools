.. tabs::

    .. code-tab:: python

        #!/usr/bin/env python
        from __future__ import annotations
        import time
        import sys
        from typing import Any

        from ooodev.adapter.lang.event_events import EventEvents
        from ooodev.events.args.event_args import EventArgs
        from ooodev.events.lo_events import Events
        from ooodev.events.lo_named_event import LoNamedEvent
        from ooodev.office.calc import Calc
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo


        class DocMonitor:
            def __init__(self) -> None:
                self.closed = False
                self.bridge_disposed = False
                loader = Lo.load_office(Lo.ConnectPipe(), opt=Lo.Options(verbose=True))

                self.doc = Calc.create_doc(loader=loader)

                self._fn_on_disposed = self.on_disposed

                self.events = Events(source=self)
                self.events.on(LoNamedEvent.BRIDGE_DISPOSED, self._fn_on_disposed)

                self._bridge_events = EventEvents()
                self._bridge_events.add_event_disposing(self._fn_on_disposing)
                Lo.bridge.addEventListener(self._bridge_events.events_listener_event)

                GUI.set_visible(True, self.doc)

            def on_disposed(self, source: Any, event_args: EventArgs) -> None:
                # just another way of knowing when bridge is gone.
                print("LO: Office bridge has gone!!")
                self.bridge_disposed = True


        def main_loop() -> None:
            dw = DocMonitor()

            # check an see if user passed in a auto terminate option
            if len(sys.argv) > 1:
                if str(sys.argv[1]).casefold() in ("t", "true", "y", "yes"):
                    Lo.delay(5000)
                    Lo.close_office()

            # while Writer is open, keep running the script unless specifically ended by user
            while 1:
                if dw.bridge_disposed is True:
                    print("\nExiting due to office bridge is gone\n")
                    raise SystemExit(1)
                time.sleep(0.1)


        if __name__ == "__main__":
            print("Press 'ctl+c' to exit script early.")
            try:
                main_loop()
            except SystemExit as e:
                SystemExit(e.code)
            except KeyboardInterrupt:
                # ctrl+c exist the script early
                print("\nExiting by user request.\n", file=sys.stderr)
                SystemExit(0)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None
