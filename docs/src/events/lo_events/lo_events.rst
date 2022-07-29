Class LoEvents
==============


``LoEvents`` is a singleton class.

Because this is a singleton class it will continue to fire events as long as your running code has one reference to ``LoEvents``.

Think of ``LoEvents`` as being global scoped.

Events can be removed via the :py:meth:`~.events.lo_events.LoEvents.remove` method.
It is not possible to remove |odev| built in events.

Perhaps a better solution to subscribe to event is the :py:class:`~.events.lo_events.Events` class as it has a local scope.

.. warning::

    Subscribing to events on this class can have unexpected side effects.
    Such as subscribed events being triggered when you thought you code was finished running.

    In most case the it is recommended to use :ref:`events_lo_events_Events` to hook events.

Each time |odev| internally runs a print command a event is raised with :py:class:`~.events.args.cancel_event_args.CancelEventArgs` as the args parameter.
If that event is hooked then the print can be canceled.

In the following example all |odev| internal print commands are canceled.

.. collapse:: Example
    :open:

    .. code-block:: python
        :emphasize-lines: 26,27,35

        #!/usr/bin/env python
        # coding: utf-8
        from __future__ import annotations
        import argparse
        from typing import Any, cast

        from ooodev.utils.lo import Lo
        from ooodev.office.write import Write
        from ooodev.utils.info import Info
        from ooodev.wrapper.break_context import BreakContext
        from ooodev.events.gbl_named_event import GblNamedEvent
        from ooodev.events.args.cancel_event_args import CancelEventArgs
        from ooodev.events.lo_events import LoEvents


        def args_add(parser: argparse.ArgumentParser) -> None:
            parser.add_argument(
                "-f",
                "--file",
                help="File path of input file to convert",
                action="store",
                dest="file_path",
                required=True,
            )

        def on_lo_print(source: Any, e: CancelEventArgs) -> None:
            e.cancel = True

        def main() -> int:
            parser = argparse.ArgumentParser(description="main")
            args_add(parser=parser)
            args = parser.parse_args()

            # hook ooodev internal printing event
            LoEvents().on(GblNamedEvent.PRINTING, on_lo_print)

            with BreakContext(Lo.Loader(Lo.ConnectSocket(headless=True))) as loader:

                fnm = cast(str, args.file_path)

                try:
                    doc = Lo.open_doc(fnm=fnm, loader=loader)
                except Exception:
                    print(f"Could not open '{fnm}'")
                    raise BreakContext.Break

                if Info.is_doc_type(obj=doc, doc_type=Lo.Service.WRITER):
                    text_doc = Write.get_text_doc(doc=doc)
                    cursor = Write.get_cursor(text_doc)
                    text = Write.get_all_text(cursor)
                    print("Text Content".center(50, "-"))
                    print(text)
                    print("-" * 50)
                else:
                    print("Extraction unsupported for this doc type")
                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            raise SystemExit(main())



.. seealso:: 

    :ref:`events_lo_events_Events`

.. include:: ../../../resources/events/event_internal_note.rst

.. autoclass:: ooodev.events.lo_events.LoEvents
    :members:
    :inherited-members: