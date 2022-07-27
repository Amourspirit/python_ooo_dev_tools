.. tabs::

    .. code-tab:: python

        #!/usr/bin/env python
        # coding: utf-8
        from __future__ import annotations
        import time
        import sys
        import types
        from typing import Any

        from ooodev.utils.lo import Lo
        from ooodev.utils.gui import GUI
        from ooodev.office.calc import Calc
        from ooodev.events.args.event_args import EventArgs
        from ooodev.events.lo_events import Events
        from ooodev.events.lo_named_event import LoNamedEvent


        class DocMonitor:
            def __init__(self) -> None:
                super().__init__()
                self.bridge_disposed = False
                loader = Lo.load_office(Lo.ConnectPipe())

                self.events = Events(source=self)
                self.events.on(LoNamedEvent.BRIDGE_DISPOSED, DocMonitor.on_disposed)

                self.doc = Calc.create_doc(loader=loader)

                GUI.set_visible(True, self.doc)

            @staticmethod
            def on_disposed(source: Any, event: EventArgs) -> None:
                print("LO: Office bridge has gone!!")
                event.event_source.bridge_disposed = True


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
                sys.exit(e.code)
            except KeyboardInterrupt:
                # ctrl+c exitst the script earily
                print("\nExiting by user request.\n", file=sys.stderr)
                sys.exit(0)