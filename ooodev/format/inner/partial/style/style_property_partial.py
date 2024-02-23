from __future__ import annotations
from typing import Any, Dict

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import event_ctx
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.props_named_event import PropsNamedEvent
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps


class StylePropertyPartial:
    """
    Partial class for Applying a Style Name to a component.
    """

    def __init__(self, component: Any, property_name: str, property_default="") -> None:
        self.__component = component
        self.__property_name = property_name
        self.__property_default = property_default

    def style_by_name(self, name: str = "") -> None:
        """
        Assign a style by name to the component.

        Args:
            name (str, optional): The name of the style to apply. If not provided, the default style is applied.

        Raises:
            CancelEventError: If the event ``before_style_by_name`` is cancelled and not handled.

        Returns:
            None:
        """

        comp = self.__component
        prop_name = self.__property_name
        cancel_set_prop = False
        cargs = None
        prop_val = name
        prop_default = self.__property_default
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_by_name.__qualname__)
            event_data: Dict[str, Any] = {
                "name": prop_val,
                "prop_name": prop_name,
                "this_component": comp,
                "cancel_set_prop": cancel_set_prop,
            }
            cargs.event_data = event_data
            self.trigger_event(StyleNameEvent.STYLE_NAME_APPLYING, cargs)
            self.trigger_event("before_style_by_name", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return None
                cargs.set("initial_event", "before_style_by_name")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style Position has been cancelled.")
                else:
                    return None
            prop_val = cargs.event_data.get("name", prop_val)
            prop_name = cargs.event_data.get("prop_name", prop_name)
            cancel_set_prop = cargs.event_data.get("cancel_set_prop", cancel_set_prop)
            comp = cargs.event_data.get("this_component", comp)

        def _on_props_setting(
            source: Any, event_args: KeyValCancelArgs, *args, **kwargs  # pylint: disable=unused-argument
        ) -> None:
            nonlocal prop_default
            # if the event property value is the default value, then trigger the default_style_by_name_prop_setting event.
            # this can bubble up to the default_style_by_name_prop_setting event and give an opportunity to change have the default property set by Props.set method.
            if event_args.value == prop_default and isinstance(self, EventsPartial):
                default_args = KeyValCancelArgs(
                    self.style_by_name.__qualname__, key=event_args.key, value=prop_default
                )
                # pylint: disable=no-member
                self.trigger_event("style_by_name_default_prop_setting", default_args)  # type: ignore
                event_args.value = default_args.value
                event_args.cancel = default_args.cancel
                event_args.handled = default_args.handled
                event_args.key = default_args.key
                event_args.value = default_args.value
                event_args.default = default_args.default

        if cancel_set_prop is False:
            if not prop_name:
                return
            if not prop_val:
                prop_val = self.__property_default

            # Subscribe to property setting event while name is being set.
            with event_ctx(
                (PropsNamedEvent.PROP_SETTING, _on_props_setting),
                source=self,
                lo_observe=True,
            ):
                mProps.Props.set(comp, **{prop_name: name})

        if cargs is not None:
            # pylint: disable=no-member
            event_args = EventArgs.from_args(cargs)
            self.trigger_event("after_style_by_name", event_args)  # type: ignore
            self.trigger_event(StyleNameEvent.STYLE_NAME_APPLIED, event_args)  # type: ignore

    def style_by_name_get(self) -> str:
        """
        Get the style name of the component.

        Returns:
            str: The name of the style.
        """
        comp = self.__component
        prop_name = self.__property_name
        cancel_set_prop = False
        cargs = None
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_by_name.__qualname__)
            event_data: Dict[str, Any] = {
                "prop_name": prop_name,
                "this_component": comp,
                "cancel_set_prop": cancel_set_prop,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_by_name_get", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return ""
                cargs.set("initial_event", "before_style_by_name_get")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style Position has been cancelled.")
                else:
                    return ""
            prop_name = cargs.event_data.get("prop_name", prop_name)
            cancel_set_prop = cargs.event_data.get("cancel_set_prop", cancel_set_prop)
            comp = cargs.event_data.get("this_component", comp)

        return mProps.Props.get(comp, prop_name, "") if prop_name else ""
