from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Type
import contextlib
import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XChild
from com.sun.star.container import XNamed

from ooo.dyn.form.form_component_type import FormComponentType

# from ooodev.adapter.lang.event_events import EventEvents
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
from ooodev.units.unit_mm import UnitMM
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.data_type.generic_unit_point import GenericUnitPoint
from ooodev.utils.data_type.generic_unit_size import GenericUnitSize
from ooodev.utils.kind.form_component_kind import FormComponentKind
from ooodev.utils.kind.language_kind import LanguageKind
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import ControlShape  # service
    from com.sun.star.awt import XControl
    from com.sun.star.awt import UnoControlModel  # service
    from com.sun.star.awt import UnoControl  # service
    from com.sun.star.uno import XInterface
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.form.forms import Forms
    from ooodev.proto.style_obj import StyleT


class FormCtlBase(
    LoInstPropsPartial,
    TheDictionaryPartial,
    PropPartial,
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
    """Base class for all form controls"""

    # both view and Model implement XServiceInfo and has supportsService() method .

    # region init
    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to ``None``.

        Returns:
            None:

        Note:
            If the :ref:`LoContext <ooodev.utils.context.lo_context.LoContext>` manager is use before this class is instantiated,
            then the Lo instance will be set using the current Lo instance. That the context manager has set.
            Generally speaking this means that there is no need to set ``lo_inst`` when instantiating this class.

        See Also:
            :ref:`ooodev.form.Forms`.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TheDictionaryPartial.__init__(self)
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        self._set_control(ctl)
        # in some cases the control model is removed from the control.
        # This means that self.get_control().getModel() return None.
        # By capturing the model here, we can set it back to the control if it is removed.
        # The model is removed for instance when a control is on a spreadsheet and the sheet is deactivated.
        self.__model = self.get_control().getModel()
        PropPartial.__init__(self, component=self.__model, lo_inst=self.lo_inst)
        trigger_args = self._get_generic_args()
        FocusEvents.__init__(self, trigger_args=trigger_args, cb=self._on_focus_listener_add_remove)
        KeyEvents.__init__(self, trigger_args=trigger_args, cb=self._on_key_events_listener_add_remove)
        MouseEvents.__init__(self, trigger_args=trigger_args, cb=self._on_mouse_listener_add_remove)
        MouseMotionEvents.__init__(self, trigger_args=trigger_args, cb=self._on_mouse_motion_listener_add_remove)
        PaintEvents.__init__(self, trigger_args=trigger_args, cb=self._on_paint_listener_add_remove)
        WindowEvents.__init__(self, trigger_args=trigger_args, cb=self._on_window_event_listener_add_remove)
        PropertyChangeImplement.__init__(self, component=self.__model, trigger_args=trigger_args)  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self.__model, trigger_args=trigger_args)  # type: ignore
        VetoableChangeImplement.__init__(self, component=self.__model, trigger_args=trigger_args)  # type: ignore
        self.__control_shape = cast("ControlShape", None)

    # endregion init

    # region Lazy Listeners

    def _on_mouse_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addMouseListener(self.events_listener_mouse)
        event.remove_callback = True

    def _on_mouse_motion_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addMouseMotionListener(self.events_listener_mouse_motion)
        event.remove_callback = True

    def _on_focus_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addFocusListener(self.events_listener_focus)
        event.remove_callback = True

    def _on_key_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addKeyListener(self.events_listener_key)
        event.remove_callback = True

    def _on_paint_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addPaintListener(self.events_listener_paint)
        event.remove_callback = True

    def _on_window_event_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addWindowListener(self.events_listener_window)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region PropPartial Overrides
    def set_property(self, **kwargs: Any) -> None:
        """
        Set property value

        Args:
            **kwargs: Variable length Key value pairs used to set properties.

        .. versionadded:: 0.39.1
        """
        ctl = self.get_control()
        started_in_design_mode = ctl.isDesignMode()
        if not started_in_design_mode:
            ctl.setDesignMode(True)
        super().set_property(**kwargs)
        if not started_in_design_mode:
            ctl.setDesignMode(False)

    # end PropPartial Overrides

    # region context manage
    def __enter__(self) -> Any:
        # new in 0.39.1
        ctl = self.get_control()
        if not ctl.isDesignMode():
            ctl.setDesignMode(True)
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        ctl = self.get_control()
        if ctl.isDesignMode():
            ctl.setDesignMode(False)

    # endregion context manage

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, FormCtlBase):
            return NotImplemented
        return self.get_control().getModel() == other.get_control().getModel()

    # region other methods

    def set_design_mode(self, on: bool) -> None:
        """
        Sets the design mode for use in a design editor.

        Normally the control will be painted directly without a peer.

        Args:
            on (bool): ``True`` to set design mode on; Otherwise, ``False``.

        .. versionadded:: 0.47.6
        """
        self.get_control().setDesignMode(on)

    def assign_script(
        self,
        interface_name: str | XInterface,
        method_name: str,
        script_name: str,
        loc: str,
        language: str | LanguageKind = LanguageKind.PYTHON,
        auto_remove_existing: bool = True,
    ) -> None:
        """
        Binds a macro to a form control.

        |lo_safe|

        Args:
            interface_name (str, XInterface): Interface Name or a UNO object that implements the ``XInterface``.
            method_name (str): Method Name.
            script_name (str): Script Name.
            loc (str): can be user, share, document, and extensions.
            language (str | LanguageKind, optional): Language. Defaults to LanguageKind.PYTHON.
            auto_remove_existing (bool, optional): Remove existing script. Defaults to ``True``.

        Returns:
            None:

        See Also:
            - `Scripting Framework URI Specification <https://wiki.openoffice.org/wiki/Documentation/DevGuide/Scripting/Scripting_Framework_URI_Specification>`_
            - :py:meth:`~.remove_script`

        .. versionchanged:: 0.47.6
            added auto_remove_existing parameter.
        """
        props = self.get_property_set()
        self._forms_class.assign_script(
            ctl_props=props,
            interface_name=interface_name,
            method_name=method_name,
            script_name=script_name,
            loc=loc,
            language=language,
            auto_remove_existing=auto_remove_existing,
        )

    def remove_script(self, interface_name: str | XInterface, method_name: str, remove_params: str = "") -> None:
        """
        Removes a script from a form control.

        Args:
            ctl_props (XPropertySet): _description_
            interface_name (str | XInterface): _description_
            method_name (str): _description_
            remove_params (str, optional): _description_. Defaults to "".

        Raises:
            RemoveScriptError: if there is an error removing the script.

        Returns:
            None:

        See Also:
            - :py:meth:`~.assign_script`

        .. versionadded:: 0.47.6
        """
        props = self.get_property_set()
        self._forms_class.remove_script(
            ctl_props=props, interface_name=interface_name, method_name=method_name, remove_params=remove_params
        )

    def get_id(self) -> int:
        """
        Gets class id for this control.

        Returns:
            int: Class Id if found, Otherwise ``-1``
        """
        props = self.get_property_set()
        with contextlib.suppress(Exception):
            return int(props.getPropertyValue("ClassId"))
        mLo.Lo.print("No class ID found for form component")
        return -1

    def get_model(self) -> UnoControlModel:
        """Gets the model for this control"""
        # upon testing in Calc, I noticed that the model is not available
        # after the sheet that the control is on is deactivated ( another sheet is activated).
        # So, I am going to try to get the model from the form that the control is on.

        ctl = self.get_control()
        model = ctl.getModel()
        if model is not None:
            return cast("UnoControlModel", model)
        if self.__model is not None:
            ctl.setModel(self.__model)
        model = ctl.getModel()
        if model is not None:
            return cast("UnoControlModel", model)
        # if we get here, then we have no model
        raise ValueError(
            "Unable to get model for control. Consider setting the model manually using add_event_listener_no_model() event."
        )

    def get_view(self) -> UnoControl:
        """Gets the view of this control"""
        return cast("UnoControl", self.get_control().getView())

    def get_control(self) -> XControl:
        """Gets the control"""
        return self.__control

    def _set_control(self, ctl: XControl) -> None:
        self.__control = ctl

    def _get_generic_args(self) -> GenericArgs:
        try:
            return self.__generic_args
        except AttributeError:
            self.__generic_args = GenericArgs(control_src=self)
            return self.__generic_args

    def get_property_set(self) -> XPropertySet:
        """
        Gets the property set for this control.
        """
        return mLo.Lo.qi(XPropertySet, self.get_model(), True)

    def apply_styles(self, *styles: StyleT) -> None:
        """
        Applies styles to control

        Args:
            *styles: Styles to apply
        """
        model = self.get_model()
        with LoContext(self.lo_inst):
            for style in styles:
                style.apply(model)

    def get_form_name(self) -> str:
        """
        Gets form name for the current control

        Args:
            ctl_model (XControlModel): control model

        Returns:
            str: form name on success; Otherwise, empty string.
        """
        child = mLo.Lo.qi(XChild, self.get_model())
        if not child:
            return ""
        parent = child.getParent()
        if not parent:
            return ""
        named = mLo.Lo.qi(XNamed, parent)
        if not named:
            return ""
        return named.getName()

    # endregion other methods

    # region Overrides
    def get_uno_srv_name(self) -> str:
        """Get Uno service name"""
        return self.get_form_component_kind().to_namespace()

    def _get_tab_index(self) -> int:
        """Gets the tab index"""
        return self.get_model().TabIndex

    def _set_tab_index(self, value: int) -> None:
        """Sets the tab index"""
        self.get_model().TabIndex = value

    # endregion Overrides

    # region Abstract Methods
    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        raise NotImplementedError("Must override get_form_component_kind")

    # endregion Abstract Methods

    # region Properties

    @property
    def component_type(self) -> int:
        """
        Gets the form component type.

        The return value is a ``com.sun.star.form.FormComponentType`` constant.

        Returns:
            int: Form component type

        .. versionadded:: 0.14.1
        """
        form_id = self.get_id()
        if form_id == -1:
            return FormComponentType.CONTROL
        return form_id

    @property
    def name(self) -> str:
        """Gets the name for the control model"""
        return self.get_model().Name

    @property
    def tab_index(self) -> int:
        """Gets/Sets the tab index"""
        return self._get_tab_index()

    @tab_index.setter
    def tab_index(self, value: int) -> None:
        self._set_tab_index(value)

    @property
    def tag(self) -> str:
        """Gets/Sets the tag"""
        return self.get_model().Tag

    @tag.setter
    def tag(self, value: str) -> None:
        self.get_model().Tag = value

    @property
    def size(self) -> GenericUnitSize[UnitMM, float]:
        """Gets the size of the control in ``UnitMM`` Values."""
        if self.control_shape is None:
            raise ValueError("control_shape property is None")
        sz = self.control_shape.getSize()
        return GenericUnitSize(UnitMM.from_mm100(sz.Width), UnitMM.from_mm100(sz.Height))

    @property
    def position(self) -> GenericUnitPoint[UnitMM, float]:
        """Gets the Position of the control in ``UnitMM`` Values."""
        if self.control_shape is None:
            raise ValueError("control_shape property is None")
        ps = self.control_shape.getPosition()
        return GenericUnitPoint(UnitMM.from_mm100(ps.X), UnitMM.from_mm100(ps.Y))

    @property
    def control_shape(self) -> ControlShape:
        """Gets the owner of the control"""
        return self.__control_shape

    @control_shape.setter
    def control_shape(self, value: ControlShape) -> None:
        self.__control_shape = value

    @property
    def _forms_class(self) -> Type[Forms]:
        """Gets the class name for the form"""
        # delay import to avoid circular import.
        try:
            return self._forms_class_instance
        except AttributeError:
            # pylint: disable=import-outside-toplevel
            from ooodev.form.forms import Forms as OooDevForms

            self._forms_class_instance = OooDevForms
        return self._forms_class_instance

    # endregion Properties
