# region imports
from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import
import unohelper
from com.sun.star.uno import XInterface
from com.sun.star.beans import XPropertySet
from com.sun.star.awt import PosSize

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.awt.focus_events import FocusEvents
from ooodev.adapter.awt.key_events import KeyEvents
from ooodev.adapter.awt.mouse_events import MouseEvents
from ooodev.adapter.awt.mouse_motion_events import MouseMotionEvents
from ooodev.adapter.awt.paint_events import PaintEvents
from ooodev.adapter.awt.window_events import WindowEvents
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.loader import lo as mLo
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.utils.partial.view_prop_partial import ViewPropPartial
from ooodev.adapter.awt.uno_control_dialog_element_partial import UnoControlDialogElementPartial
from ooodev.units.unit_px import UnitPX
from ooodev.units.unit_app_font_height import UnitAppFontHeight
from ooodev.units.unit_app_font_width import UnitAppFontWidth
from ooodev.units.unit_app_font_x import UnitAppFontX
from ooodev.units.unit_app_font_y import UnitAppFontY
from ooodev.utils.kind.point_size_kind import PointSizeKind
from ooodev.utils.builder.default_builder import DefaultBuilder


if TYPE_CHECKING:
    from com.sun.star.awt import XControlModel
    from com.sun.star.awt import XControl
    from com.sun.star.awt import UnoControlDialogElement  # service
    from com.sun.star.beans import XMultiPropertySet
    from com.sun.star.awt import XWindowPeer
    from com.sun.star.awt import XWindow
    from ooodev.proto.style_obj import StyleT
    from ooodev.units.unit_obj import UnitT
# endregion imports

# pylint: disable=unused-argument

# Model Position and Size are in AppFont units. View Size and Position are in Pixel units.


class CtlBase(unohelper.Base, LoInstPropsPartial, ViewPropPartial, ModelPropPartial, EventsPartial):
    """Control Base Class"""

    def __new__(cls, ctl: Any, *args, **kwargs):
        return super().__new__(cls, *args, **kwargs)

    # region Dunder Methods
    def __init__(self, ctl: Any) -> None:
        unohelper.Base.__init__(self)
        LoInstPropsPartial.__init__(self, lo_inst=mLo.Lo.current_lo)
        EventsPartial.__init__(self)
        model = ctl.getModel()
        ViewPropPartial.__init__(self, obj=ctl)
        ModelPropPartial.__init__(self, obj=model)
        self._set_control(ctl)

    def _set_control(self, ctl: Any) -> None:
        self._ctl_view = ctl

    # endregion Dunder Methods

    # region other methods

    def _get_generic_args(self) -> GenericArgs:
        try:
            return self.__generic_args
        except AttributeError:
            self.__generic_args = GenericArgs(control_src=self)
            return self.__generic_args

    def get_view(self) -> XControl:
        return self._ctl_view

    get_view_ctl = get_view

    def get_uno_srv_name(self) -> str:
        """Get Uno service name"""
        raise NotImplementedError

    def get_control_props(self) -> XPropertySet | None:
        """
        Gets property set for a control model

        Args:
            control_model (Any): control model

        Returns:
            XPropertySet | None: Property set
        """
        try:
            return mLo.Lo.qi(XPropertySet, self.get_model())
        except Exception:
            return None

    # endregion other methods

    # region Other Methods
    def get_model(self) -> XControlModel:
        """Gets the Model for the control"""
        return self.get_view().getModel()

    # endregion Other Methods

    # region Properties

    @property
    def visible(self) -> bool:
        """Gets/Sets the visible state for the control"""
        try:
            model = cast(Any, self.get_view().getModel())
            return model.EnableVisible
        except Exception:
            return True

    @visible.setter
    def visible(self, value: bool) -> None:
        with contextlib.suppress(Exception):
            model = cast(Any, self.get_model())
            model.EnableVisible = value

    # endregion Properties


# region Control Listener Base
class _UnoControlDialogElementOverrides:

    def __new__(cls, *args, **kwargs):
        return super().__new__(cls, *args, **kwargs)

    def __init__(self, ctl: Any):
        """
        Constructor
        """
        self.__model = cast("UnoControlDialogElement", ctl.getModel())

    # region UnoControlDialogElementPartial Overrides
    @property
    def width(self) -> UnitPX:
        """Gets the width of the control"""
        model = self.__model
        return UnitPX.from_app_font(model.Width, PointSizeKind.WIDTH)
        # view = mLo.Lo.qi(XView, self.get_view(), True)
        # return UnitPX(view.getSize().Width)

    @width.setter
    def width(self, value: int | UnitT) -> None:
        """
        Gets/Sets the width of the control in Pixel units.

        When setting can be an integer in ``Pixels`` Units or a ``UnitT``.

        Returns:
            UnitPX: Width of the control.

        Note:
            The Width is in Pixel units; however, the model is in AppFont units.
            This property will convert the AppFont units to Pixel units.
            When set using ``UnitAppFontWidth`` no conversion is done.
        """
        model = self.__model
        if isinstance(value, UnitAppFontWidth):
            model.Width = int(value)
            return
        val = UnitPX.from_unit_val(value)
        model.Width = round(val.get_value_app_font(PointSizeKind.WIDTH))

    @property
    def height(self) -> UnitPX:
        """
        Gets/Sets the height of the control in Pixel units.

        When setting can be an integer in ``Pixels`` Units or a ``UnitT``.

        Returns:
            UnitPX: Height of the control.

        Note:
            The Height is in Pixel units; however, the model is in AppFont units.
            This property will convert the AppFont units to Pixel units.
            When set using ``UnitAppFontHeight`` no conversion is done.
        """
        model = self.__model
        return UnitPX.from_app_font(model.Height, PointSizeKind.HEIGHT)

    @height.setter
    def height(self, value: int | UnitT) -> None:
        model = self.__model
        if isinstance(value, UnitAppFontHeight):
            model.Height = int(value)
            return
        val = UnitPX.from_unit_val(value)
        model.Height = round(val.get_value_app_font(PointSizeKind.HEIGHT))

    # endregion UnoControlDialogElementPartial Overrides


class CtlListenerBase(
    CtlBase,
    # UnoControlDialogElementPartial,
    # _UnoControlDialogElementOverrides,
    FocusEvents,
    KeyEvents,
    MouseEvents,
    MouseMotionEvents,
    PaintEvents,
    WindowEvents,
    PropertyChangeImplement,
    PropertiesChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing table ConfigurationProvider Component.

    Note:
        This is a Dynamic class that is created at runtime.
        This means that the class is created at runtime and not defined in the source code.
        In addition, the class may be created with additional classes implemented.

        The Type hints for this class at design time may not be accurate.
        To check if a class implements a specific interface, use the ``isinstance`` function
        or :py:meth:`~.InterfacePartial.is_supported_interface` methods which is always available in this class.
    """

    # pylint: disable=unused-argument

    def __new__(cls, *args, **kwargs):

        # Because the Dialog class also inherits this class the extra check for ctl is needed.

        ctl = kwargs.get("ctl", None)
        if ctl is None:
            try:
                ctl = args[0]
            except IndexError:
                ctl = None
        if ctl is None:
            return super().__new__(cls, *args, **kwargs)
        builder = _get_control_listener_builder(ctl)
        clz = builder.get_class_type(
            name="ooodev.dialog.dl_control.ctl_base.CtlListenerBase", base_class=cls, set_mod_name=False
        )
        return super().__new__(clz, *args, **kwargs)

    # region Dunder Methods
    def __init__(self, ctl: Any) -> None:
        CtlBase.__init__(self, ctl)
        if isinstance(self, UnoControlDialogElementPartial):
            UnoControlDialogElementPartial.__init__(self, self.get_model())  # type: ignore
            _UnoControlDialogElementOverrides.__init__(self, ctl)  # type: ignore
        generic_args = self._get_generic_args()  # type: ignore
        # The Events callback methods are invoked when any event is added or removed.
        FocusEvents.__init__(self, trigger_args=generic_args, cb=self._on_focus_listener_add_remove)
        KeyEvents.__init__(self, trigger_args=generic_args, cb=self._on_key_events_listener_add_remove)
        MouseEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_listener_add_remove)
        MouseMotionEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_motion_listener_add_remove)
        PaintEvents.__init__(self, trigger_args=generic_args, cb=self._on_paint_listener_add_remove)
        WindowEvents.__init__(self, trigger_args=generic_args, cb=self._on_window_event_listener_add_remove)
        model = self.get_model()  # type: ignore
        PropertyChangeImplement.__init__(self, component=cast(XPropertySet, model), trigger_args=generic_args)
        PropertiesChangeImplement.__init__(self, component=cast("XMultiPropertySet", model), trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=cast(XPropertySet, model), trigger_args=generic_args)

    # endregion Dunder Methods
    # region UnoControlDialogElementPartial properties
    if TYPE_CHECKING:
        # The properties from this point are optional. The builder will add them if needed.
        # There are include here for type hinting only.
        @property
        def height(self) -> UnitPX:
            """
            Gets/Sets the height of the control in Pixel units.

            When setting can be an integer in ``Pixels`` Units or a ``UnitT``.

            Returns:
                UnitPX: Height of the control.

            Note:
                The Height is in Pixel units; however, the model is in AppFont units.
                This property will convert the AppFont units to Pixel units.
                When set using ``UnitAppFontHeight`` no conversion is done.
            """
            ...

        @height.setter
        def height(self, value: int | UnitT) -> None: ...

        @property
        def name(self) -> str:
            """
            Gets/Sets the name of the control.
            """
            ...

        @name.setter
        def name(self, value: str) -> None: ...

        @property
        def x(self) -> UnitAppFontX:
            """
            Gets/Sets the horizontal position of the control.

            When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

            Returns:
                UnitAppFontX: Horizontal position of the control.
            """
            ...

        @x.setter
        def x(self, value: int | UnitT) -> None: ...

        @property
        def y(self) -> UnitAppFontY:
            """
            Gets/Sets the vertical position of the control.

            When setting can be an integer in ``AppFont`` Units or a ``UnitT``.

            Returns:
                UnitAppFontY: Vertical position of the control.
            """
            ...

        @y.setter
        def y(self, value: int | UnitT) -> None: ...

        @property
        def step(self) -> int:
            """
            Gets/Sets the step of the control.
            """
            ...

        @step.setter
        def step(self, value: int) -> None: ...
        @property
        def tab_index(self) -> int:
            """
            Gets/Sets the tab index of the control.
            """
            ...

        @tab_index.setter
        def tab_index(self, value: int) -> None: ...

        @property
        def tag(self) -> str:
            """
            Gets/Sets the tag of the control.
            """
            ...

        @tag.setter
        def tag(self, value: str) -> None: ...

        @property
        def width(self) -> UnitPX:
            """Gets the width of the control"""
            ...

        @width.setter
        def width(self, value: int | UnitT) -> None:
            """
            Gets/Sets the width of the control in Pixel units.

            When setting can be an integer in ``Pixels`` Units or a ``UnitT``.

            Returns:
                UnitPX: Width of the control.

            Note:
                The Width is in Pixel units; however, the model is in AppFont units.
                This property will convert the AppFont units to Pixel units.
                When set using ``UnitAppFontWidth`` no conversion is done.
            """
            ...

    # endregion UnoControlDialogElementPartial properties

    # region Overrides
    def get_uno_srv_name(self) -> str:
        """Get Uno service name"""
        raise NotImplementedError

    # endregion Overrides

    # region Lazy Listeners

    # Listeners such as mouse, mouse motion, focus, key, paint, and window are added lazily.
    # Some listeners such as mouse motion may be expensive.
    # Each time the mouse moves, the listener is invoked.
    # By lazy loading listeners are only added when needed.
    # For example:
    #     ctl_button_ok.add_event_mouse_entered(on_mouse_entered)
    # This would call _on_mouse_listener_add_remove() below, which would add the mouse listener to the class in a lazy manor.

    def _on_mouse_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addMouseListener(self.events_listener_mouse)
        event.remove_callback = True

    def _on_mouse_motion_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addMouseMotionListener(self.events_listener_mouse_motion)
        event.remove_callback = True

    def _on_focus_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addFocusListener(self.events_listener_focus)
        event.remove_callback = True

    def _on_key_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addKeyListener(self.events_listener_key)
        event.remove_callback = True

    def _on_paint_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addPaintListener(self.events_listener_paint)
        event.remove_callback = True

    def _on_window_event_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addWindowListener(self.events_listener_window)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region other methods
    def get_property_set(self) -> XPropertySet:
        """Gets the property set for this control"""
        return mLo.Lo.qi(XPropertySet, self.get_model(), True)

    # endregion other methods


def _get_control_listener_builder(ctl: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    model = ctl.getModel()
    builder = DefaultBuilder(ctl)
    if hasattr(model, "Name"):
        builder.add_import(
            "ooodev.adapter.awt.uno_control_dialog_element_partial.UnoControlDialogElementPartial",
            optional=False,
            init_kind=1,
            check_kind=0,
        )
        builder.add_import(
            "ooodev.dialog.dl_control.ctl_base._UnoControlDialogElementOverrides",
            optional=False,
            init_kind=1,
            check_kind=0,
        )
    return builder


# endregion Control Listener Base

# region DialogControlBase


class DialogControlBase(CtlListenerBase):
    """Dialog Control Base Class. Only for Controls that have a model that can be added to a dialog"""

    def __new__(cls, *args, **kwargs):
        ctl = kwargs.get("ctl", None)
        if ctl is None:
            ctl = args[0]
        builder = _get_dialog_control_base_builder(ctl)
        clz = builder.get_class_type(
            name="ooodev.dialog.dl_control.ctl_base.DialogControlBase", base_class=cls, set_mod_name=False
        )
        return super().__new__(clz, *args, **kwargs)

    def __init__(self, ctl: Any) -> None:
        CtlListenerBase.__init__(self, ctl)
        if isinstance(self, _DialogControlBaseUnoControlElements):
            _DialogControlBaseUnoControlElements.__init__(self, ctl)

    # region Overrides
    def get_uno_srv_name(self) -> str:
        """Get Uno service name"""
        raise NotImplementedError

    # endregion Overrides

    # region other methods
    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind"""
        return DialogControlKind.from_value(self.get_uno_srv_name())

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind"""
        with contextlib.suppress(Exception):
            kind = self.get_control_kind()
            return DialogControlNamedKind.from_str(kind.name)
        return DialogControlNamedKind.UNKNOWN

    def apply_styles(self, *styles: StyleT) -> None:
        """
        Applies styles to control

        Args:
            *styles: Styles to apply
        """
        model = self.get_model()
        for style in styles:
            style.apply(model)

    # endregion other methods

    # region Properties
    @property
    def tip_text(self) -> str:
        """Gets/Sets the tip text"""
        with contextlib.suppress(Exception):
            model = cast(Any, self.get_model())
            return model.HelpText
        return ""

    @tip_text.setter
    def tip_text(self, value: str) -> None:
        with contextlib.suppress(Exception):
            model = cast(Any, self.get_model())
            model.HelpText = value

    # useful alias
    help_text = tip_text
    if TYPE_CHECKING:
        # The properties from this point are optional. The builder will add them if needed.
        # There are include here for type hinting only.
        @property
        def tab_index(self) -> int:
            """Gets/Sets the tab index"""
            ...

        @tab_index.setter
        def tab_index(self, value: int) -> None: ...

        @property
        def step(self) -> int:
            """Gets/Sets the step"""
            ...

        @step.setter
        def step(self, value: int) -> None: ...

        @property
        def tag(self) -> str:
            """Gets/Sets the tag"""
            ...

        @tag.setter
        def tag(self, value: str) -> None: ...

    # endregion Properties


class _DialogControlBaseUnoControlElements:

    def __new__(cls, *args, **kwargs):
        return super().__new__(cls, *args, **kwargs)

    def __init__(self, ctl: Any):
        """
        Constructor
        """
        self.__model = cast("UnoControlDialogElement", ctl.getModel())

    # region UnoControlDialogElementPartial Overrides
    @property
    def tab_index(self) -> int:
        """Gets/Sets the tab index"""
        with contextlib.suppress(Exception):
            return self.__model.TabIndex
        return -1

    @tab_index.setter
    def tab_index(self, value: int) -> None:
        with contextlib.suppress(Exception):
            self.__model.TabIndex = value

    @property
    def step(self) -> int:
        """Gets/Sets the step"""
        with contextlib.suppress(Exception):
            return self.__model.Step
        return 0

    @step.setter
    def step(self, value: int) -> None:
        with contextlib.suppress(Exception):
            self.__model.Step = value

    @property
    def tag(self) -> str:
        """Gets/Sets the tag"""
        with contextlib.suppress(Exception):
            return self.__model.Tag
        return ""

    @tag.setter
    def tag(self, value: str) -> None:
        with contextlib.suppress(Exception):
            self.__model.Tag = value


def _get_dialog_control_base_builder(ctl: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    model = ctl.getModel()
    builder = DefaultBuilder(ctl)
    if hasattr(model, "Name"):
        builder.add_import(
            "ooodev.dialog.dl_control.ctl_base._DialogControlBaseUnoControlElements",
            optional=False,
            init_kind=1,
            check_kind=0,
        )
    return builder


# endregion DialogControlBase


def _create_control(model_name: str, win: XWindowPeer, **kwargs) -> Any:
    """
    Creates a new instance of the control.

    Returns:
        CtlFixedText: New instance of the control
    """

    def set_size_pos(
        ctl: XWindow, x: int | UnitT = -1, y: int | UnitT = -1, width: int | UnitT = -1, height: int | UnitT = -1
    ) -> None:
        """
        Set Position and size for a control.

        Args:
            ctl (XWindow): Control that implements XWindow
            x (int, UnitT, optional): X Position. Defaults to -1.
            y (int, UnitT, optional): Y Position. Defaults to -1.
            width (int, UnitT, optional): Width. Defaults to -1.
            height (int, UnitT, optional): Height. Defaults to -1.
        """
        px_x = int(UnitPX.from_unit_val(x))
        px_y = int(UnitPX.from_unit_val(y))
        px_width = int(UnitPX.from_unit_val(width))
        px_height = int(UnitPX.from_unit_val(height))
        if px_x < 0 and px_y < 0 and px_width < 0 and px_height < 0:
            return

        pos_size = None
        if px_x > -1 and px_y > -1 and px_width > -1 and px_height > -1:
            pos_size = PosSize.POSSIZE
        elif px_x > -1 and px_y > -1:
            pos_size = PosSize.POS
        elif px_width > -1 and px_height > -1:
            pos_size = PosSize.SIZE
        if pos_size is not None:
            ctl.setPosSize(px_x, px_y, px_width, px_height, pos_size)

    x = kwargs.pop("x", -1)
    y = kwargs.pop("y", -1)
    width = kwargs.pop("width", -1)
    height = kwargs.pop("height", -1)
    lo_inst = kwargs.pop("lo_inst", mLo.Lo.current_lo)

    model = cast(Any, lo_inst.create_instance_mcf(XInterface, model_name))
    ctrl = cast(Any, lo_inst.create_instance_mcf(XInterface, model.DefaultControl))
    props = cast(XPropertySet, lo_inst.qi(XPropertySet, model))
    if props is not None:
        prop_info = props.getPropertySetInfo()

        for k, v in kwargs.items():
            if prop_info.hasPropertyByName(k):
                props.setPropertyValue(k, v)
    ctrl.setModel(model)
    set_size_pos(ctrl, x, y, width, height)
    ctrl.createPeer(win.Toolkit, win)  # type: ignore

    return ctrl
