from __future__ import annotations

from typing import Any, cast, TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.form.submission.submission_veto_listener import SubmissionVetoListener

if TYPE_CHECKING:
    from com.sun.star.form.submission import XSubmission
    from ooodev.utils.type_var import EventArgsCallbackT, CancelEventArgsCallbackT, ListenerEventCallbackT


class SubmissionVetoEvents:
    """
    Class for managing Submission Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: SubmissionVetoListener | None = None,
        subscriber: XSubmission | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (SubmissionVetoListener | None, optional): Listener that is used to manage events.
            subscriber (XSubmission, optional): An UNO object that implements the ``XSubmission`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addSubmissionVetoListener(self.__listener)
        else:
            self.__listener = SubmissionVetoListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_submitting(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Is invoked when a component, at which the listener has been registered, is about to submit its data.
        If event is canceled and the cancel args are not handled then a ``VetoException`` will be raised.

        The callback ``CancelEventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.

        Note:
            When callback event is invoked it will contain a :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`
            instance as the event. The ``CancelEventArgs.cancel`` can be set to ``True`` to cancel the submission.
            Also if canceled the ``CancelEventArgs.handled`` can be set to ``True`` to indicate that the submission
            should be performed. The ``CancelEventArgs.event_data`` will contain the original ``com.sun.star.lang.EventObject``
            that triggered the update.

            Also the ``CancelEventArgs`` can set a ``message`` value that will be used as the message for the ``VetoException``.

            If the ``event.set("skip_veto_exception", True)`` is set then the ``VetoException`` will not be raised.
            This is probably not a good idea but it is there if you need it.

            The following example shows how to use the ``CancelEventArgs`` to cancel the submission of data.

            .. code-block:: python

                def on_submitting(src: Any, event: CancelEventArgs, *args: Any, **kwargs: Any) -> None:
                    if not validate_data():
                        event.cancel = True
                        event.set("message", "Canceling due to data validation fail.")
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="submitting")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("submitting", cb)

    def add_event_submission_veto_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the broadcaster is about to be disposed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """

        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("disposing", cb)

    def remove_event_submitting(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="submitting", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("submitting", cb)

    def remove_event_submission_veto_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("disposing", cb)

    @property
    def events_listener_submission_veto(self) -> SubmissionVetoListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events


def on_lazy_cb(source: Any, event: ListenerEventArgs) -> None:
    """
    Callback that is invoked when an event is added or removed.

    This method is generally used to add the listener to the component in a lazy manner.
    This means this callback will only be called once in the lifetime of the component.

    Args:
        source (Any): Expected to be an instance of SubmissionVetoEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, SubmissionVetoEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XSubmission", source.component)  # type: ignore
    comp.addSubmissionVetoListener(source.events_listener_submission_veto)
    event.remove_callback = True
