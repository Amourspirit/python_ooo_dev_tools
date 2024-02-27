from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.form.submission import XSubmissionVetoListener

from ooo.dyn.util.veto_exception import VetoException

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.form.submission import XSubmission


class SubmissionVetoListener(AdapterBase, XSubmissionVetoListener):
    """
    Is implement by components which want to observe (and probably veto) the submission of data.

    See Also:
        `API XSubmissionVetoListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1form_1_1submission_1_1XSubmissionVetoListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XSubmission | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XSubmission, optional): An UNO object that implements the ``com.sun.star.form.XSubmission`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)

        if subscriber:
            subscriber.addSubmissionVetoListener(self)

    def submitting(self, event: EventObject) -> None:
        """
        Is invoked when a component, at which the listener has been registered, is about to submit its data.

        If event is canceled and the cancel args are not handled then a ``VetoException`` will be raised.

        Args:
            event (EventObject): Event data for the event.

        Raises:
            com.sun.star.util.VetoException: ``VetoException``

        Returns:
            None:

        Note:
            When ``submitting`` event is invoked it will contain a :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`
            instance as the trigger event. When the event is triggered the ``CancelEventArgs.cancel`` can be set to ``True``
            to cancel the submission. Also if canceled the ``CancelEventArgs.handled`` can be set to ``True`` to indicate that the submission
            should be performed. The ``CancelEventArgs.event_data`` will contain the original ``com.sun.star.lang.EventObject``
            that triggered the update.

            Also the ``CancelEventArgs`` can set a ``message`` value that will be used as the message for the ``VetoException``.

            If the ``event.set("skip_veto_exception", True)`` is set then the ``VetoException`` will not be raised.
            This is probably not a good idea but it is there if you need it.

            The following example shows how to use the ``CancelEventArgs`` to cancel the submission of data.

            .. code-block:: python

                def on_submitting(src: Any, event: CancelEventArgs) -> None:
                    if not validate_data():
                        event.cancel = True
                        event.set("message", "Canceling due to data validation fail.")
        """

        # raise VetoException("VetoException", self, 0)
        cancel_args = CancelEventArgs(self.__class__.__qualname__)
        cancel_args.event_data = event
        cancel_args.set("message", "VetoException Raise due to CancelEventArgs")
        self._trigger_direct_event("submitting", cancel_args)
        if cancel_args.cancel:
            if not cancel_args.handled:
                # just in case raise a VetoException here is not the correct thing to do
                # then give an out.
                skip_veto_exception = cancel_args.get("skip_veto_exception", False)
                if skip_veto_exception:
                    return
                # if the cancel event was handled then we return True to indicate that the submit should be performed
                msg = cancel_args.get("message", "VetoException Raise due to CancelEventArgs")
                raise VetoException(msg, self)

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
