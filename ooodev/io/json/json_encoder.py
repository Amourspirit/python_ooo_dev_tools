from __future__ import annotations
from typing import Any
import json
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.gen_util import NULL_OBJ
from ooodev.utils.helper.dot_dict import DotDict


class JsonEncoder(json.JSONEncoder):
    """
    General JSON Encoder class.

    This class can be used to encode objects to JSON and is generally used as a base class for other classes that need to encode objects to JSON.

    Any object that has a ``to_json()`` method will be encoded using that method.

    See :py:class:`ooodev.gui.menu.popup.PopupCreator` for an example of a class that uses this class.
    """

    def on_json_encode(self, obj: Any) -> Any:
        """
        Protected method to encode object to JSON.
        Can be overridden by subclasses.

        Args:
            obj (Any): Object to encode. This is the object that json is currently encoding.

        Returns:
            Any: The result of the encoding. The default is ``NULL_OBJ`` which means that the encoding is not handled.
        """
        return NULL_OBJ

    def default(self, obj: Any) -> Any:
        """
        JsonEncoder default method.

        Args:
            obj (Any): Data to be encoded.

        Returns:
            Any: Encoded data.

        Note:
            This method or the ``on_json_encode()`` can be overridden by subclasses to encode objects to JSON.
            If this class is a subclass of ``EventsPartial``, the ``json_encoding`` event is triggered before encoding.
            The event data is a dictionary with the key ``obj`` containing the object to be encoded.
            If the event data ``result`` key is set then the value is returned as the result of the encoding.
        """
        if isinstance(self, EventsPartial):
            eargs = EventArgs(self)
            eargs.event_data = DotDict(obj=obj)
            self.trigger_event("json_encoding", eargs)
            if "result" in eargs.event_data:
                return eargs.event_data["result"]
        if hasattr(obj, "to_json"):
            return obj.to_json()
        result = self.on_json_encode(obj)
        if result is not NULL_OBJ:
            return result
        return super().default(obj)
