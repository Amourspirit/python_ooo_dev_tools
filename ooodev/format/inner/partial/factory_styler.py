from __future__ import annotations
from typing import Any, TYPE_CHECKING, Callable

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.utils.context.lo_context import LoContext
from ooodev.format.inner.partial.factory_name_base import FactoryNameBase

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
else:
    LoInst = Any


class FactoryStyler(FactoryNameBase):
    """
    Class for Line Properties.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        super().__init__(factory_name, component, lo_inst)
        self.before_event_name = "before_style_border_line"
        self.after_event_name = "after_style_border_line"

    def style(self, factory: Callable[[str], Any], **kwargs) -> Any:
        """
        Style Font.

        Raises:
            CancelEventError: If the ``before_*`` event is cancelled and not handled.

        Returns:
            Any: Style instance or ``None`` if cancelled.
        """
        comp = self._component
        factory_name = self._factory_name
        has_events = False
        cargs = None
        event_data = kwargs.copy()
        cancel_apply = False
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style.__qualname__)
            event_data["factory_name"] = factory_name
            event_data["this_component"] = comp
            event_data["cancel_apply"] = cancel_apply

            cargs.event_data = event_data
            self.trigger_event(self.before_event_name, cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", self.before_event_name)
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style has been cancelled.")
                else:
                    return None
            comp = cargs.event_data.pop("this_component", comp)
            factory_name = cargs.event_data.pop("factory_name", factory_name)
            cancel_apply = cargs.event_data.pop("cancel_apply", cancel_apply)
            event_data = cargs.event_data

        styler = factory(factory_name)
        fe = styler(**event_data)
        # fe.factory_name = factory_name

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        if not cancel_apply:
            with LoContext(self._lo_inst):
                fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event(self.after_event_name, EventArgs.from_args(cargs))  # type: ignore
        return fe

    def style_get(
        self,
        factory: Callable[[str], Any],
        call_method_name: str = "from_obj",
        event_name_suffix: str = "_get",
        obj_arg_name: str = "obj",
        **kwargs,
    ) -> Any:
        """
        Gets the Style.

        Args:
            factory (Callable[[str], Any]): Factory function.
            call_method_name (str): Method name to call. Defaults to "from_obj".
            event_name_suffix (str): Event name suffix. Defaults to "_get".
            obj_arg_name (str): Object argument name. Defaults to "obj". If empty, the object is not passed to the method.
            kwargs: Keyword arguments to pass to factory method.

        Raises:
            CancelEventError: If the event is cancelled and not handled.

        Returns:
            TransparencyT | None: Area transparency style or ``None`` if cancelled.
        """

        # general method for calling style methods.
        # see ooodev.format.inner.partial.chart2.numbers.numbers_numbers_partial.NumbersNumbersPartial
        # for an example of how to use this method.
        def call_method(style_obj, obj, method_name, **kwargs):
            method = getattr(style_obj, method_name, None)
            if method is not None and callable(method):
                if obj is not None:
                    kwargs[obj_arg_name] = obj
                return method(**kwargs)
            else:
                raise AttributeError(f"Method {method_name} not found.")

        event_name = f"{self.before_event_name}{event_name_suffix}"
        comp = self._component
        factory_name = self._factory_name
        cargs = None
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_get.__qualname__)
            event_data = {
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event(event_name, cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", event_name)
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style get has been cancelled.")
                else:
                    return None
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = factory(factory_name)
        obj = kwargs.pop("obj", comp)
        if not obj_arg_name:
            obj = None
        try:
            # style = styler.from_obj(obj=comp, **kwargs)
            style = call_method(style_obj=styler, obj=obj, method_name=call_method_name, **kwargs)
        except mEx.DisabledMethodError:
            return None

        style.set_update_obj(comp)
        return style
