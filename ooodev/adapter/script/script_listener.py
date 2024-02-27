from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.script import XScriptListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.cancel_event_args import CancelEventArgs

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.script import ScriptEvent
    from com.sun.star.script import XEventAttacherManager


class ScriptListener(AdapterBase, XScriptListener):
    """
    Specifies a listener which is to be notified about state changes in a grid control

    **since**

        OOo 3.1

    See Also:
        `API XScriptListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1form_1_1XGridControlListener.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XEventAttacherManager | None = None
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XEventAttacherManager, optional): An UNO object that implements the ``com.sun.star.form.XEventAttacherManager`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)

        if subscriber:
            subscriber.addScriptListener(self)

    def approveFiring(self, event: ScriptEvent) -> bool:
        """
        Event is invoked when a ``vetoable event`` occurs at the object.
        If event is canceled then the firing will be canceled.

        Args:
            event (EventObject): Event data for the event.

        Returns:
            bool: ``True`` if the firing should be performed, ``False`` otherwise.

        Note:
            When ``approveFiring`` event is invoked it will contain a :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`
            instance as the trigger event. When the event is triggered the ``CancelEventArgs.cancel`` can be set to ``True``
            to cancel the firing. Also if canceled the ``CancelEventArgs.handled`` can be set to ``True`` to indicate that the firing
            should be performed. The ``CancelEventArgs.event_data`` will contain the original ``com.sun.star.lang.EventObject``
            that triggered the update.
        """
        cancel_args = CancelEventArgs(self.__class__.__qualname__)
        cancel_args.event_data = event
        self._trigger_direct_event("approveFiring", cancel_args)
        if cancel_args.cancel:
            if CancelEventArgs.handled:
                # if the cancel event was handled then we return True to indicate that the update should be performed
                return True
            return False
        return True

    def firing(self, event: ScriptEvent) -> None:
        """
        Event is invoked when an event takes place.

        For that a ``ScriptEventDescriptor`` is registered at and attached to an object by an ``XEventAttacherManager``.

        Args:
            event (ScriptEvent): Event data for the event.

        Returns:
            None:
        """
        self._trigger_event("firing", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.

        Args:
            event (EventObject): Event data for the event.

        Returns:
            None:
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", event)
