from __future__ import annotations
from typing import Any, TYPE_CHECKING
import contextlib
import uno
from com.sun.star.awt import XItemListener
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import ItemEvent


class ItemListener(AdapterBase, XItemListener):
    """
    Makes it possible to receive events from a component when the state of an item changes.

    See Also:
        `API XItemListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XItemListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: Any = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (Any, optional): An UNO object that has a ``addItemListener()`` Method.
                If passed in then this listener instance is automatically added to it.
                Valid objects are: RadioButton, ComboBox, CheckBox,
                XItemEventBroadcaster or any other UNO object that has ``addItemListener()`` method.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            # several object such as Listbox, Combobox, Button, etc. can add an ActionListener.
            # There is no common interface for this, so we have to try them all.
            with contextlib.suppress(AttributeError):
                subscriber.addItemListener(self)

    # region XItemListener
    def itemStateChanged(self, event: ItemEvent) -> None:
        """
        Event is invoked when an item changes its state.
        """
        self._trigger_event("itemStateChanged", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", event)

    # endregion XItemListener
