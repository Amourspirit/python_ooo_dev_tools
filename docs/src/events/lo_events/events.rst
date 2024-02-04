.. _events_lo_events_Events:

Class Events
============


``Events`` class is a locally scoped class.

Once an instance of ``Events`` is created it can be used to subscribe to any internal event of |odev|.

.. collapse:: Example
    :open:

    In the following example ``local_events`` is created in the ``open_remove_save()`` method.

    While ``open_remove_save()`` is running any row index listed in ``protected_rows`` will not
    be deleted even if they are passed into ``open_remove_save()``

    As soon as the  ``open_remove_save()`` method is executed ``local_events`` goes out of scope.
    When ``local_events`` is out of scope the events it was subscribe to get released.
    Also setting ``local_events = None`` would release any events it was subscribed to.

    .. tabs::

        .. code-tab:: python

            from typing import Any
            from ooodev.loader.lo import Lo
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

``Events`` uses weak reference internally.
For this reason assigning events in a class instance may not work as expected.

For instance assigning a class method as an event handler will not work unless the class method is a static method.

Sometimes it may be useful to pass a class instance to the event so a class property can be set.
The way ``Events`` enables class instance to be sent to a method is by way of ``EventArgs.event_source`` property.

.. collapse:: Example
    :open:

    In the following example class constructor creates ``Events`` instance and assigns to class instance.
    It is important that the ``Events`` instance be assigned to class or it will go out of scope when constructor is done
    and that will result in no events being triggered for class.

    | When ``Events`` is constructed it passes in the current class instance
    | ``self.events = Events(source=self)``
    | This allows the instance to be accessed when event is triggered.

    Note that ``on_disposed()`` is a static method. ``Events`` is not able to attach to instance methods.

    The object passed into ``Events`` constructor (class instance in this case) are assigned to the ``EventArg.event_source`` property.
    Now class instance properties can be set when the event is triggered ``event.event_source.bridge_disposed = True``.

    For a complete example see `Office Window Monitor <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_monitor>`_

    .. include:: ../../../resources/events/events_in_class_ex.rst

.. include:: ../../../resources/events/event_internal_note.rst

.. seealso::

    :ref:`Chapter 4 - Dispatching <ch04_dispatching>`

.. autoclass:: ooodev.events.lo_events.Events
    :members:
    :inherited-members: