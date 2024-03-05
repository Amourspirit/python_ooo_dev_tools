from __future__ import annotations
from typing import Any, cast, Dict, TYPE_CHECKING, Callable

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.exceptions import ex as mEx
from ooodev.utils.context.lo_context import LoContext
from ooodev.format.inner.partial.factory_name_base import FactoryNameBase

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.style_t import StyleT
else:
    LoInst = Any
    StyleT = Any


class FactoryStyler(FactoryNameBase):
    """
    Class for Generic Styler.

    This class is generally used as a helper to create partial classes that can style components,
    such as, ``ooodev.format.inner.partial.font.font_only_partial.FontOnlyPartial``.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor.

        Args:
            factory_name (str): The name that the factory will use to get a style object.
            component (Any): The component the style will be applied to.
            lo_inst (LoInst | None, optional): Loader instance. Defaults to None. Used in multi-document environments.
        """
        super().__init__(factory_name, component, lo_inst)
        # self.before_event_name = "before_style_border_line"
        # self.after_event_name = "after_style_border_line"

    def style(self, factory: Callable[[str], Any], **kwargs: Any) -> Any:
        """
        Style Font.

        Raises:
            CancelEventError: If the ``before_*`` event is cancelled and not handled.

        Returns:
            Any: Style instance or ``None`` if cancelled.
        """
        comp = self._component
        factory_name = self._factory_name
        cancel_apply = False

        cargs = self._pre_style(**kwargs)
        if cargs is None:
            return None
        event_data = cast(Dict[str, Any], cargs.event_data.copy())
        comp = event_data.pop("this_component", comp)
        factory_name = event_data.pop("factory_name", factory_name)
        cancel_apply = event_data.pop("cancel_apply", cancel_apply)
        fe = event_data.pop("styler_object", None)

        styler = factory(factory_name)
        if fe is None:
            fe = styler(**event_data)
        # fe.factory_name = factory_name

        fe.add_event_observer(self.event_observer)  # type: ignore
        backup = False
        if not cancel_apply:
            event_data["factory_name"] = factory_name
            event_data["this_component"] = comp
            event_data["styler_object"] = fe

            c_backup_args = self._style_backup(fe, event_data)
            backup = not c_backup_args.cancel

            with LoContext(self._lo_inst):
                fe.apply(comp)

            if backup:
                _ = self._style_restore(style=fe, event_data=event_data, c_backup_args=c_backup_args)

        fe.set_update_obj(comp)
        if cancel_apply is False:
            event_data["styler_object"] = fe
            event_data["this_component"] = comp
            self._post_style(event_data)
        return fe

    def _pre_style(self, **kwargs) -> CancelEventArgs | None:
        comp = self._component
        factory_name = self._factory_name
        event_data = kwargs.copy()
        cancel_apply = False
        cargs = CancelEventArgs(self._pre_style.__qualname__)
        event_data["factory_name"] = factory_name
        event_data["this_component"] = comp
        event_data["cancel_apply"] = cancel_apply
        event_data["styler_object"] = None

        cargs.event_data = event_data
        self.trigger_event(StyleNameEvent.STYLE_APPLYING, cargs)
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
        return cargs

    def _post_style(self, event_data: Dict[str, Any]):
        event_args = EventArgs(source=self._post_style.__qualname__)
        event_args.event_data = event_data
        self.trigger_event(self.after_event_name, EventArgs.from_args(event_args))  # type: ignore
        self.trigger_event(StyleNameEvent.STYLE_APPLIED, EventArgs.from_args(event_args))  # type: ignore

    def _style_backup(self, style: StyleT, event_data: dict[str, Any]) -> CancelEventArgs:
        comp = event_data["this_component"]
        c_backup_args = CancelEventArgs(source=self._style_backup.__qualname__, cancel=True)

        c_backup_args.event_data = event_data
        self.trigger_event(StyleNameEvent.STYLE_BACKING_UP, c_backup_args)
        self.trigger_event(f"{self.before_event_name}_backup", c_backup_args)
        backup = not c_backup_args.cancel
        if backup:
            with LoContext(self._lo_inst):
                style.backup(comp)
            self.trigger_event(StyleNameEvent.STYLE_BACKED_UP, EventArgs.from_args(c_backup_args))
        return c_backup_args

    def _style_restore(
        self, style: StyleT, event_data: Dict[str, Any], c_backup_args: CancelEventArgs
    ) -> CancelEventArgs:
        clear_on_restore = True
        comp = event_data["this_component"]
        event_data["clear_on_restore"] = clear_on_restore
        c_restore_args = CancelEventArgs(source=self.style.__qualname__, cancel=True)
        c_restore_args.event_data = event_data
        self.trigger_event(StyleNameEvent.STYLE_RESTORING, c_restore_args)
        self.trigger_event(f"{self.before_event_name}_restore", c_backup_args)
        if c_restore_args.cancel is False:
            clear_on_restore = c_restore_args.event_data.get("clear_on_restore", clear_on_restore)
            with LoContext(self._lo_inst):
                style.restore(comp, clear_on_restore)
            self.trigger_event(StyleNameEvent.STYLE_RESTORED, EventArgs.from_args(c_restore_args))
        return c_restore_args

    def style_apply(self, style: StyleT, **kwargs: Any) -> StyleT | None:
        """
        Applies a know style to this component.

        Args:
            style (StyleT): Style to Apply.

        Returns:
            StyleT | None: Style that was Applied or None if cancelled.
        """
        comp = self._component
        factory_name = self._factory_name
        cancel_apply = False

        cargs = self._pre_style(**kwargs)
        if cargs is None:
            return None
        event_data = cast(Dict[str, Any], cargs.event_data.copy())
        comp = event_data.pop("this_component", comp)
        factory_name = event_data.pop("factory_name", factory_name)
        cancel_apply = event_data.pop("cancel_apply", cancel_apply)

        style.remove_event_observer(self.event_observer)  # type: ignore
        style.add_event_observer(self.event_observer)  # type: ignore

        backup = False
        if not cancel_apply:
            event_data["factory_name"] = factory_name
            event_data["this_component"] = comp
            event_data["styler_object"] = style

            c_backup_args = self._style_backup(style, event_data)
            backup = not c_backup_args.cancel

            with LoContext(self._lo_inst):
                style.apply(comp)

            if backup:
                _ = self._style_restore(style=style, event_data=event_data, c_backup_args=c_backup_args)
        style.set_update_obj(comp)
        if cancel_apply is False:
            event_data["styler_object"] = style
            event_data["this_component"] = comp
            self._post_style(event_data)
        return style

    def style_get(
        self,
        factory: Callable[[str], Any],
        call_method_name: str = "from_obj",
        event_name_suffix: str = "_get",
        obj_arg_name: str = "obj",
        **kwargs: Any,
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
