from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from com.sun.star.view import XSelectionSupplier
from com.sun.star.frame import XModel

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.loader import lo as mLo
from ooodev.adapter.view.selection_change_listener import SelectionChangeListener

if TYPE_CHECKING:
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class SelectionChangeEvents:
    """
    Class for managing Selection Change Events.

    This class is usually inherited by control classes that implement ``com.sun.star.view.XSelectionChangeListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: SelectionChangeListener | None = None,
        doc: Any = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (SelectionChangeListener | None, optional): Listener that is used to manage events.
            doc (Any, Optional): Office Document. If document is passed then ``SelectionChangeListener`` instance
                is automatically added.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if doc:
                model = mLo.Lo.qi(XModel, doc)
                if model is None:
                    mLo.Lo.print("Could not get model for doc")
                    return

                supp = mLo.Lo.qi(XSelectionSupplier, model.getCurrentController())
                if supp is None:
                    mLo.Lo.print("Could not attach selection change listener")
                    return
                supp.addSelectionChangeListener(self.__listener)
        else:
            self.__listener = SelectionChangeListener(trigger_args=trigger_args, doc=doc)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_selection_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the selection changes.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        # sourcery skip: class-extract-method
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="selectionChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("selectionChanged", cb)

    def add_event_selection_change_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_selection_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="selectionChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("selectionChanged", cb)

    def remove_event_selection_change_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_selection_change(self) -> SelectionChangeListener:
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
        source (Any): Expected to be an instance of SelectionChangeEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, SelectionChangeEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast(XSelectionSupplier, source.component)  # type: ignore
    comp.addSelectionChangeListener(source.events_listener_selection_change)
    event.remove_callback = True
