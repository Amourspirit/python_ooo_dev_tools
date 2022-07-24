.. _events_lo_events_Events:

Class Events
============


``Events`` class is a locally scoped class.

Once an instance of ``Events`` is created it can be used to subscribe to any internal event of |app_name_short|.

.. collapse:: Example
    :open:

    In the following example ``local_events`` is created in the ``open_remove_save()`` method.

    While ``open_remove_save()`` is running any row index listed in ``protected_rows`` will not
    be deleted even if they are passed into ``open_remove_save()``

    As soon as the  ``open_remove_save()`` method is executed ``local_events`` goes out of scope.
    When ``local_events`` is out of scope the events it was subscribe to get released.
    Also setting ``local_events = None`` would release any events it was subscribed to.

    .. code-block:: python

        from typing import Any
        from ooodev.utils.lo import Lo
        from ooodev.office.calc import Calc
        from ooodev.events.lo_events import Events
        from ooodev.events.args.calc.sheet_cancel_args import SheetCancelArgs
        from ooodev.events.calc_named_event import CalcNamedEvent

        
        protected_rows = (1, 3, 10, 15, 18)

        def protect_row(source:Any, args:SheetCancelArgs) -> None:
            nonlocal protected_rows
            if args.index in protected_rows:
                args.cancel = True

        def open_remove_save(fnm:str, *indexes:int) -> None:
            if len(indexes) == 0:
                return

            local_events = Events()
            events.on(CalcNamedEvent.SHEET_ROW_DELETING, protect_row)
            idxs = list(indexes)
            idxs.sort()
            idxs.reverse() # must remove from hightest to lowest.

            with Lo.Loader(Lo.ConnectSocket()) as loader:
                doc = Calc.open_doc(fnm=fnm, loader=loader)
                sheet = Calc.get_sheet(doc=doc, index=0)
                for idx in idxs:
                    Calc.delete_row(sheet=sheet, idx=idx)
                
                Lo.save(doc)

.. include:: ../../../resources/events/event_interal_note.rst

.. seealso::

    :ref:`Chapter 4 - Dispatching <ch04sec05>`

.. autoclass:: ooodev.events.lo_events.Events
    :members:
    :inherited-members: