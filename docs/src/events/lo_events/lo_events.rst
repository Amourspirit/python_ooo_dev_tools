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

Each time |odev| internally runs a print command a event is raised with :py:class:`~.events.args.cancel_event_args.CancelEventArgs` as the args parameter.
If that event is hooked then the print can be canceled.

In the following example all |odev| internal print commands are canceled.

.. collapse:: Example

    .. code-block:: python
        
        from typing import Any
        from ooodev.events.lo_events import LoEvents
        from ooodev.events.args.cancel_event_args import CancelEventArgs
        from ooodev.events.gbl_named_event import GblNamedEvent

        def cancel_print(source: Any, args:CancelEventArgs):
            args.cancel = True
        
        LoEvents().on(GblNamedEvent.PRINTING, cancel_print)


.. seealso:: 

    :ref:`events_lo_events_Events`

.. include:: ../../../resources/events/event_interal_note.rst

.. autoclass:: ooodev.events.lo_events.LoEvents
    :members:
    :inherited-members: