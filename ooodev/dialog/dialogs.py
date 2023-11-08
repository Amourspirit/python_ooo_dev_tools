# coding: utf-8
# region Imports
from __future__ import annotations
import datetime
from typing import TYPE_CHECKING, Any, Iterable, Tuple, cast
from enum import IntEnum
from ..utils import lo as mLo
from ..utils.date_time_util import DateUtil
from ..utils.kind.border_kind import BorderKind as BorderKind
from ..utils.kind.orientation_kind import OrientationKind as OrientationKind
from ..utils.kind.align_kind import AlignKind as AlignKind

import uno

from com.sun.star.awt import XControl
from com.sun.star.awt import XControlContainer
from com.sun.star.awt import XControlModel
from com.sun.star.awt import XDialog
from com.sun.star.awt import XDialogProvider
from com.sun.star.awt import XToolkit
from com.sun.star.awt import XTopWindow
from com.sun.star.awt import XWindow
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNameContainer
from com.sun.star.lang import XInitialization
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.lang import XServiceInfo

# com.sun.star.awt.PushButtonType
from ooo.dyn.awt.push_button_type import PushButtonType as PushButtonType
from ooo.dyn.awt.pos_size import PosSize as PosSize
from ooo.dyn.style.vertical_alignment import VerticalAlignment as VerticalAlignment
from ooo.dyn.awt.image_scale_mode import ImageScaleModeEnum as ImageScaleModeEnum

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButton  # service
    from com.sun.star.awt import UnoControlButtonModel  # service
    from com.sun.star.awt import UnoControlCheckBox  # service
    from com.sun.star.awt import UnoControlCheckBoxModel  # service
    from com.sun.star.awt import UnoControlComboBox  # service
    from com.sun.star.awt import UnoControlComboBoxModel  # service
    from com.sun.star.awt import UnoControlCurrencyField  # service
    from com.sun.star.awt import UnoControlCurrencyFieldModel  # service
    from com.sun.star.awt import UnoControlDateField
    from com.sun.star.awt import UnoControlDateFieldModel
    from com.sun.star.awt import UnoControlEdit  # service
    from com.sun.star.awt import UnoControlEditModel  # service
    from com.sun.star.awt import UnoControlFileControl  # service
    from com.sun.star.awt import UnoControlFileControlModel  # service
    from com.sun.star.awt import UnoControlFixedHyperlink  # service
    from com.sun.star.awt import UnoControlFixedHyperlinkModel  # service
    from com.sun.star.awt import UnoControlFixedLine  # service
    from com.sun.star.awt import UnoControlFixedLineModel  # service
    from com.sun.star.awt import UnoControlFixedText  # service
    from com.sun.star.awt import UnoControlFixedTextModel  # service
    from com.sun.star.awt import UnoControlFormattedField  # service
    from com.sun.star.awt import UnoControlFormattedFieldModel  # service
    from com.sun.star.awt import UnoControlGroupBox  # service
    from com.sun.star.awt import UnoControlGroupBoxModel  # service
    from com.sun.star.awt import UnoControlImageControl  # service
    from com.sun.star.awt import UnoControlImageControlModel  # service
    from com.sun.star.awt import UnoControlListBox  # service
    from com.sun.star.awt import UnoControlListBoxModel  # service
    from com.sun.star.awt import UnoControlNumericField  # service
    from com.sun.star.awt import UnoControlNumericFieldModel  # service
    from com.sun.star.awt import UnoControlRadioButton  # service
    from com.sun.star.awt import UnoControlRadioButtonModel  # service
    from com.sun.star.awt import UnoControlPatternField  # service
    from com.sun.star.awt import UnoControlPatternFieldModel  # service
    from com.sun.star.awt import UnoControlProgressBar  # service
    from com.sun.star.awt import UnoControlProgressBarModel  # service
    from com.sun.star.awt.tab import UnoControlTabPageContainer  # service
    from com.sun.star.awt.tab import UnoControlTabPageContainerModel  # service
    from com.sun.star.awt.tab import UnoControlTabPage  # service
    from com.sun.star.awt.tab import UnoControlTabPageModel  # service
    from com.sun.star.awt import UnoControlDialog  # service
    from com.sun.star.awt import UnoControlDialogModel  # service
    from com.sun.star.container import XNameAccess
    from com.sun.star.lang import EventObject
else:
    UnoControlButton = object
    UnoControlCheckBox = object
    UnoControlComboBox = object
    UnoControlCurrencyField = object
    UnoControlDateField = object
    UnoControlEdit = object
    UnoControlFileControl = object
    UnoControlFixedHyperlink = object
    UnoControlFixedLine = object
    UnoControlFixedText = object
    UnoControlFormattedField = object
    UnoControlGroupBox = object
    UnoControlImageControl = object
    UnoControlListBox = object
    UnoControlNumericField = object
    UnoControlRadioButton = object
    UnoControlPatternField = object
    UnoControlProgressBar = object
    UnoControlTabPageContainer = object
    UnoControlTabPage = object
    UnoControlDialog = object
# endregion Imports


class Dialogs:
    class StateEnum(IntEnum):
        NOT_CHECKED = 0
        """State not checked"""
        CHECKED = 1
        """State checked"""
        DONT_KNOW = 2
        """State don't know"""

    # region    load & execute a dialog
    @staticmethod
    def load_dialog(script_name: str) -> XDialog:
        """
        Create a dialog for the given script name

        Args:
            script_name (str): script name

        Raises:
            Exception: if unable to create dialog

        Returns:
            XDialog: Dialog instance
        """
        dp = mLo.Lo.create_instance_mcf(XDialogProvider, "com.sun.star.awt.DialogProvider", raise_err=True)

        try:
            return dp.createDialog(f"vnd.sun.star.script:{script_name}?location=application")
        except Exception as e:
            mLo.Lo.print("Could not access the Dialog Provider")
            raise e

    @staticmethod
    def load_addon_dialog(extension_id: str, dialog_fnm: str) -> XDialog:
        """
        Loads addon dialog

        Args:
            extension_id (str): Addon id
            dialog_fnm (str): Addon file path

         Raises:
            Exception: if unable to create dialog

        Returns:
            XDialog: Dialog instance
        """
        dp = mLo.Lo.create_instance_mcf(XDialogProvider, "com.sun.star.awt.DialogProvider", raise_err=True)
        return dp.createDialog(f"vnd.sun.star.extension://{extension_id}/{dialog_fnm}")

    # endregion load & execute a dialog

    # region    access a control/component inside a dialog

    @staticmethod
    def find_control(dialog_ctrl: XControl, name: str) -> XControl:
        """
        Finds control by name

        Args:
            dialog_ctrl (XControl): Control
            name (str): Name to find

        Returns:
            XControl: Control
        """
        ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)
        return ctrl_con.getControl(name)

    @classmethod
    def show_control_info(cls, dialog_ctrl: XControl) -> None:
        """
        Prints info for a control to console

        Args:
            dialog_ctrl (XControl): Control

        Returns:
            None:
        """
        controls = cls.get_dialog_controls_arr(dialog_ctrl)
        print(f"No of controls: {len(controls)}")
        for i, ctl in enumerate(controls):
            print(f"{i}. Name: {cls.get_control_name(ctl)}")
            print(f"  Default Control: {cls.get_control_class_id(ctl)}")
            print()

    @staticmethod
    def get_dialog_controls_arr(dialog_ctrl: XControl) -> Tuple[XControl, ...]:
        """
        Gets all controls for a given control

        Args:
            dialog_ctrl (XControl): control

        Returns:
            Tuple[XControl, ...]: controls
        """
        ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)
        return ctrl_con.getControls()

    @staticmethod
    def get_control_props(control_model: Any) -> XPropertySet:
        """
        Gets property set for a control model

        Args:
            control_model (Any): control model

        Returns:
            XPropertySet: Property set
        """
        return mLo.Lo.qi(XPropertySet, control_model, True)

    @classmethod
    def get_control_name(cls, control: XControl) -> str:
        """
        Get the name of a control

        Args:
            control (XControl): control

        Returns:
            str: control name
        """
        props = cls.get_control_props(control.getModel())
        return str(props.getPropertyValue("Name"))

    @classmethod
    def get_control_class_id(cls, control: XControl) -> str:
        """
        Gets control class id

        Args:
            control (XControl): control

        Returns:
            str: class id
        """
        props = cls.get_control_props(control.getModel())
        return str(props.getPropertyValue("DefaultControl"))

    @classmethod
    def get_event_source_name(cls, event: EventObject) -> str:
        """
        Get event source name

        Args:
            event (EventObject): event

        Returns:
            str: event source name
        """
        return cls.get_control_name(cls.get_event_control(event))

    @staticmethod
    def get_event_control(event: EventObject) -> XControl:
        """
        Gets event control from event

        Args:
            event (EventObject): event

        Returns:
            XControl: control
        """
        return mLo.Lo.qi(XControl, event.Source, True)

    # endregion access a control/component inside a dialog

    # region    convert dialog into other forms
    @staticmethod
    def get_dialog(dialog_ctrl: XControl) -> XDialog:
        """
        Gets dialog from dialog control

        Args:
            dialog_ctrl (XControl): control

        Returns:
            XDialog: dialog
        """
        return mLo.Lo.qi(XDialog, dialog_ctrl, True)

    @staticmethod
    def get_dialog_control(dialog: XDialog) -> XControl:
        """
        Gets dialog control

        Args:
            dialog (XDialog): dialog

        Returns:
            XControl: control.
        """
        return mLo.Lo.qi(XControl, dialog, True)

    @staticmethod
    def get_dialog_window(dialog_ctrl: XControl) -> XTopWindow:
        """
        Gets dialog window

        Args:
            dialog_ctrl (XControl): dialog control

        Returns:
            XTopWindow: Top window instance
        """
        return mLo.Lo.qi(XTopWindow, dialog_ctrl, True)

    # endregion convert dialog into other forms

    # region    create a dialog
    @classmethod
    def create_dialog(
        cls,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        title: str,
        **props: Any,
    ) -> UnoControlDialog:
        """
        Creates a dialog

        Args:
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int): Height
            title (str): title
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create dialog.

        Returns:
            UnoControlDialog: Control

        See Also:
            `API UnoControlDialogModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogModel.html>`_
        """
        # sourcery skip: raise-specific-error
        try:
            dialog_ctrl = cast(
                UnoControlDialog,
                mLo.Lo.create_instance_mcf(XControl, "com.sun.star.awt.UnoControlDialog", raise_err=True),
            )
            control_model = mLo.Lo.create_instance_mcf(
                XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True
            )
            dialog_ctrl.setModel(control_model)

            ctl_props = cls.get_control_props(dialog_ctrl.getModel())

            ctl_props.setPropertyValue("Title", title)
            ctl_props.setPropertyValue("Name", "OfficeDialog")

            ctl_props.setPropertyValue("Step", 0)
            ctl_props.setPropertyValue("Moveable", True)
            ctl_props.setPropertyValue("TabIndex", 0)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            cls._set_size_pos(dialog_ctrl, x, y, width, height)
            return dialog_ctrl
        except Exception as e:
            raise Exception(f"Could not create dialog control: {e}") from e

    @classmethod
    def create_dialog_peer(cls, dialog_ctrl: XControl) -> XDialog:
        """
        Gets a dialog

        Args:
            dialog_ctrl (XControl): control

        Returns:
            XDialog: Dialog
        """
        xwindow = mLo.Lo.qi(XWindow, dialog_ctrl, True)
        # set the dialog window invisible until it is executed
        xwindow.setVisible(False)

        xtoolkit = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit", raise_err=True)
        window_parent_peer = xtoolkit.getDesktopWindow()

        dialog_ctrl.createPeer(xtoolkit, window_parent_peer)

        # dialog_component = mLo.Lo.qi(XComponent, dialog_ctrl)
        # dialog_component.dispose() # free window resources
        # commented out or the Add-on dialog crashes when called a second time
        # because createPeer() cannot find a model
        return cls.get_dialog(dialog_ctrl)

    # endregion create a dialog

    # region    add components to a dialog

    @staticmethod
    def get_dialog_nm_con(ctrl: XControl) -> XNameContainer:
        """
        Gets Name container from control

        Args:
            ctrl (XControl): Dialog control

        Returns:
            XNameContainer: Name Container
        """
        return mLo.Lo.qi(XNameContainer, ctrl.getModel(), True)

    @staticmethod
    def create_name(elem_container: XNameAccess, name: str) -> str:
        """
        Creates a name.

        Make a unique string by appending a number to the supplied name

        Args:
            elem_container (XNameAccess): container
            name (str): current name

        Returns:
            str: a name not in container.
        """
        used_name = True
        i = 1
        nm = f"{name}{i}"
        while used_name:
            used_name = elem_container.hasByName(nm)
            if used_name:
                i += 1
                nm = f"{name}{i}"
        return nm

    @classmethod
    def insert_button(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        btn_type: PushButtonType | None = None,
        name: str = "",
        **props: Any,
    ) -> UnoControlButton:
        """
        Insert Button Control

        Args:
            dialog_ctrl (XControl): control
            label (str): Button Label
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height. Defaults to 20.
            btn_type (PushButtonType | None, optional): Type of Button.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create button control

        Returns:
            XControl: Button control

        See Also:
            `API UnoControlButtonModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlButtonModel.html>`_
        """
        # sourcery skip: raise-specific-error
        if btn_type is None:
            btn_type = PushButtonType.STANDARD
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlButtonModel", msf.createInstance("com.sun.star.awt.UnoControlButtonModel"))

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "CommandButton")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Label", label)
            # ctl_props.setPropertyValue("PushButtonType", btn_type)
            uno_any = uno.Any("short", btn_type)  # type: ignore
            uno.invoke(ctl_props, "setPropertyValue", ("PushButtonType", uno_any))  # type: ignore
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlButton, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create button control: {e}") from e

    @classmethod
    def insert_check_box(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 8,
        tri_state: bool = True,
        state: Dialogs.StateEnum = StateEnum.CHECKED,
        name: str = "",
        **props: Any,
    ) -> UnoControlCheckBox:
        """
        Inserts a check box control

        Args:
            dialog_ctrl (XControl): Control
            label (str): Checkbox label text
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to ``8``.
            tri_state (StateEnum, optional): Specifies that the control may have the state "don't know". Defaults to ``StateEnum.CHECKED``.
            state (int, optional): specifies the state of the control. Defaults to ``1``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create checkbox control.

        Returns:
            UnoControlCheckBox: Check box control

        See Also:
            `API UnoControlCheckBoxModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlCheckBoxModel.html>`_
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlCheckBoxModel", msf.createInstance("com.sun.star.awt.UnoControlCheckBoxModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "CheckBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("Label", label)
            ctl_props.setPropertyValue("TriState", tri_state)
            ctl_props.setPropertyValue("State", int(state))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlCheckBox, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create check box control: {e}") from e

    @classmethod
    def insert_combo_box(
        cls,
        dialog_ctrl: XControl,
        *,
        entries: Iterable[str],
        x: int,
        y: int,
        width: int,
        height: int = 12,
        max_text_len: int = 0,
        drop_down: bool = True,
        read_only: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> UnoControlComboBox:
        """
        Insert a combo box control

        Args:
            dialog_ctrl (XControl): Control
            entries (Iterable[str]): Combo box entries
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to 12.
            max_text_len (int, optional): Specifies the maximum character count, There's no limitation, if set to 0. Defaults to 0.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to True.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to False.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create combo box control

        Returns:
            UnoControlComboBox: Combo box control

        See Also:
            `API UnoControlComboBoxModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlComboBoxModel.html>`_
        """
        # sourcery skip: raise-specific-error
        try:
            max_text_len = max(max_text_len, 0)
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlComboBoxModel", msf.createInstance("com.sun.star.awt.UnoControlComboBoxModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "ComboBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("Dropdown", drop_down)
            ctl_props.setPropertyValue("MaxTextLen", max_text_len)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("ReadOnly", read_only)
            if entries:
                uno_strings = uno.Any("[]string", tuple(entries))  # type: ignore
                uno.invoke(ctl_props, "setPropertyValue", ("StringItemList", uno_strings))  # type: ignore
                # ctl_props.setPropertyValue("StringItemList", tuple(entries))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlComboBox, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create combo box control: {e}") from e

    @classmethod
    def insert_currency_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 12,
        value: float = 0.0,
        min: float = -1000000.0,
        max: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> UnoControlCurrencyField:
        """
        Inserts a currency control

        Args:
            dialog_ctrl (XControl): Control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to ``12``.
            value (float, optional): Control Value. Defaults to 0.0.
            min (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create currency field box control

        Returns:
            UnoControlCurrencyField: Currency field control.
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlCurrencyFieldModel", msf.createInstance("com.sun.star.awt.UnoControlCurrencyFieldModel")
            )
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "CurrencyField")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("ValueMin", min)
            ctl_props.setPropertyValue("ValueMax", max)
            ctl_props.setPropertyValue("ValueStep", increment)
            ctl_props.setPropertyValue("Spin", spin_button)
            ctl_props.setPropertyValue("Value", value)
            ctl_props.setPropertyValue("DecimalAccuracy", accuracy)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlCurrencyField, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create text field control: {e}") from e

    @classmethod
    def insert_date_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        date_value: datetime.datetime | None = None,
        min_date: datetime.datetime = datetime.datetime(1900, 1, 1, 0, 0, 0, 0),
        max_date: datetime.datetime = datetime.datetime(2200, 12, 31, 0, 0, 0, 0),
        drop_down: bool = True,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> UnoControlDateField:
        """
        Create a new control of type DateField in the actual dialog.

        Args:
            dialog_ctrl (XControl): Control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to ``20``.
            date_value (datetime.datetime | None, optional): _description_. Defaults to None.
            min_date (datetime.datetime, optional): _description_. Defaults to datetime.datetime(1900, 1, 1, 0, 0, 0, 0).
            max_date (datetime.datetime, optional): _description_. Defaults to datetime.datetime(2200, 12, 31, 0, 0, 0, 0).
            drop_down (bool, optional): Specifies if the control is a dropdown. Defaults to True.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create date field control

        Returns:
            UnoControlDateField: Date field control
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlDateFieldModel", msf.createInstance("com.sun.star.awt.UnoControlDateFieldModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "DateField")

            # set properties in the model
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Dropdown", drop_down)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("DateMin", DateUtil.date_to_uno_date(min_date))
            ctl_props.setPropertyValue("DateMax", DateUtil.date_to_uno_date(max_date))
            if date_value is not None:
                ctl_props.setPropertyValue("Date", DateUtil.date_to_uno_date(date_value))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlDateField, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create text field control: {e}") from e

    @classmethod
    def insert_file_control(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> UnoControlFileControl:
        """
        Create a new control of type FileControl in the actual dialog

        Args:
            dialog_ctrl (XControl): control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height. Defaults to ``20``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create file control control

        Returns:
            UnoControlFileControl: File Control
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlFileControlModel", msf.createInstance("com.sun.star.awt.UnoControlFileControlModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "FileControl")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlFileControl, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create file control: {e}") from e

    @classmethod
    def insert_fixed_line(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 1,
        orientation: OrientationKind = OrientationKind.HORIZONTAL,
        name: str = "",
        **props: Any,
    ) -> UnoControlFixedLine:
        """
        Create a new control of type Fixed Line in the actual dialog

        Args:
            dialog_ctrl (XControl): control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height. Defaults to ``1``.
            orientation (OrientationKind, optional): Orientation. Defaults to ``OrientationKind.HORIZONTAL``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create file control control

        Returns:
            UnoControlFixedLine: Fixed Line Control
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlFixedLineModel", msf.createInstance("com.sun.star.awt.UnoControlFixedLineModel"))

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "FixedLine")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Orientation", int(orientation))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlFixedLine, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create fixed line control: {e}") from e

    @classmethod
    def insert_formatted_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        value: float | None = None,
        min: float = -1000000.0,
        max: float = 1000000.0,
        spin_button: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> UnoControlFormattedField:
        """
        Create a new control of type FormattedField in the actual dialog.

        Args:
            dialog_ctrl (XControl): Control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to ``12``.
            value (float, optional): Control Value. Defaults to 0.0.
            min (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create formatted field control control

        Returns:
            UnoControlFileControl: File Control
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlFormattedFieldModel", msf.createInstance("com.sun.star.awt.UnoControlFormattedFieldModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "FormattedField")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("EffectiveMin", min)
            ctl_props.setPropertyValue("EffectiveMax", max)
            ctl_props.setPropertyValue("Spin", spin_button)
            ctl_props.setPropertyValue("Border", int(border))
            if not value is None:
                ctl_props.setPropertyValue("EffectiveValue", value)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlFormattedField, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create formatted field control control: {e}") from e

    @classmethod
    def insert_group_box(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        label: str = "",
        name: str = "",
        **props: Any,
    ) -> UnoControlGroupBox:
        """
        Create a new control of type GroupBox in the actual dialog.

        Args:
            dialog_ctrl (XControl): control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create group box control control

        Returns:
            UnoControlGroupBox: Group box Control
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlGroupBoxModel", msf.createInstance("com.sun.star.awt.UnoControlGroupBoxModel"))

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "GroupBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            if label:
                ctl_props.setPropertyValue("Label", label)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlGroupBox, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create Group box control: {e}") from e

    @classmethod
    def insert_hyperlink(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        url: str,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        align: AlignKind = AlignKind.LEFT,
        vert_align: VerticalAlignment = VerticalAlignment.TOP,
        multiline: bool = False,
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        **props: Any,
    ) -> UnoControlFixedHyperlink:
        """
        Create a new control of type Hyperlink in the actual dialog.

        Args:
            dialog_ctrl (XControl): control
            label (str): Hyperlink label
            url (str): Hyperlink URL
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height. Defaults to ``20``.
            align (AlignKind, optional): Horizontal alignment. Defaults to ``AlignKind.LEFT``.
            vert_align (VerticalAlignment, optional): Vertical alignment. Defaults to ``VerticalAlignment.TOP``.
            multiline (bool, optional): Specifies if the control can display multiple lines of text. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create group box control control

        Returns:
            UnoControlFixedHyperlink: Group box Control
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlFixedHyperlinkModel", msf.createInstance("com.sun.star.awt.UnoControlFixedHyperlinkModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "Hyperlink")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Align", align.value)
            ctl_props.setPropertyValue("VerticalAlign", vert_align)
            ctl_props.setPropertyValue("MultiLine", multiline)
            ctl_props.setPropertyValue("Border", int(border))

            if label:
                ctl_props.setPropertyValue("Label", label)
            if url:
                ctl_props.setPropertyValue("URL", url)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlFixedHyperlink, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create Group box control: {e}") from e

    @classmethod
    def insert_image_control(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        border: BorderKind = BorderKind.BORDER_3D,
        scale: int | ImageScaleModeEnum = ImageScaleModeEnum.NONE,
        image_url: str = "",
        name: str = "",
        **props: Any,
    ) -> UnoControlImageControl:
        """
        Create a new control of type ImageControl in the actual dialog.

        Args:
            dialog_ctrl (XControl): control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height. Defaults to ``20``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create image control

        Returns:
            UnoControlImageControl: Image Control
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlImageControlModel", msf.createInstance("com.sun.star.awt.UnoControlImageControlModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "ImageControl")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Border", int(border))

            scale = int(scale)
            if scale != ImageScaleModeEnum.NONE.value:
                ctl_props.setPropertyValue("ScaleImage", True)
                ctl_props.setPropertyValue("ScaleMode", scale)
            else:
                ctl_props.setPropertyValue("ScaleImage", False)

            if image_url:
                ctl_props.setPropertyValue("ImageURL", image_url)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlImageControl, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create file control: {e}") from e

    @classmethod
    def insert_label(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        name: str = "",
        **props: Any,
    ) -> UnoControlFixedText:
        """
        Insert a label into a control

        Args:
            dialog_ctrl (XControl): Control
            label (str): Contents of label
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Default ``20``
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create label

        Returns:
            UnoControlFixedText: control

        See Also:
            `API UnoControlFixedTextModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlFixedTextModel.html>`_
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)

            model = cast("UnoControlFixedTextModel", msf.createInstance("com.sun.star.awt.UnoControlFixedTextModel"))

            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "FixedText")

            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Label", label)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)
            result = cast(UnoControlFixedText, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create fixed text control: {e}") from e

    @classmethod
    def insert_list_box(
        cls,
        dialog_ctrl: XControl,
        *,
        entries: Iterable[str],
        x: int,
        y: int,
        width: int,
        height: int = 100,
        drop_down: bool = False,
        read_only: bool = False,
        line_count: int = 5,
        multi_select: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> UnoControlListBox:
        """
        Insert a combo box control

        Args:
            dialog_ctrl (XControl): Control
            entries (Iterable[str]): Combo box entries
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to 40.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to True.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to False.
            line_count (int, optional): Specifies the number of lines to display. Defaults to 5.
            multi_select (int, optional): Specifies if multiple entries can be selected. Defaults to False.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create list box control

        Returns:
            UnoControlListBox: List box control
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlListBoxModel", msf.createInstance("com.sun.star.awt.UnoControlListBoxModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "ListBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("LineCount", line_count)
            ctl_props.setPropertyValue("Dropdown", drop_down)
            ctl_props.setPropertyValue("MultiSelection", multi_select)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("ReadOnly", read_only)
            if entries:
                uno_strings = uno.Any("[]string", tuple(entries))  # type: ignore
                uno.invoke(ctl_props, "setPropertyValue", ("StringItemList", uno_strings))  # type: ignore
                # ctl_props.setPropertyValue("StringItemList", tuple(entries))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlListBox, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create list box control: {e}") from e

    @classmethod
    def insert_password_field(
        cls,
        dialog_ctrl: XControl,
        *,
        text: str,
        x: int,
        y: int,
        width: int,
        height: int = 12,
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        **props: Any,
    ) -> UnoControlEdit:
        """
        Inserts a password field.

        Args:
            dialog_ctrl (XControl): Control
            text (str): Text value
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to ``12``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create text field

        Returns:
            UnoControlEdit: Text Field Control

        See Also:
            :py:meth:`~.dialogs.Dialogs.insert_text_field`
        """
        return cls.insert_text_field(
            dialog_ctrl=dialog_ctrl,
            text=text,
            x=x,
            y=y,
            width=width,
            height=height,
            echo_char="*",
            border=border,
            name=name,
            **props,
        )

    @classmethod
    def insert_pattern_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        edit_mask: str = "",
        literal_mask: str = "",
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> UnoControlPatternField:
        """
        Create a new control of type PatternField in the actual dialog.

        Args:
            dialog_ctrl (XControl): control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height.
            edit_mask (str, optional): Specifies a character code that determines what the user may enter. Defaults to ``""``.
            literal_mask (str, optional): Specifies the initial values that are displayed in the pattern field. Defaults to ``""``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create pattern field control control

        Returns:
            UnoControlPatternField: Pattern Field Control
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlPatternFieldModel", msf.createInstance("com.sun.star.awt.UnoControlPatternFieldModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "PatternField")
            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("EditMask", edit_mask)
            ctl_props.setPropertyValue("LiteralMask", literal_mask)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlPatternField, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create numeric field control: {e}") from e

    @classmethod
    def insert_numeric_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        value: float | None = None,
        min: float = -1000000.0,
        max: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> UnoControlNumericField:
        """
        Create a new control of type GroupBox in the actual dialog.

        Args:
            dialog_ctrl (XControl): control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height.
            value (float, optional): Control Value. Defaults to 0.0.
            min (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create numeric field control control

        Returns:
            UnoControlNumericField: Group box Control
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlNumericFieldModel", msf.createInstance("com.sun.star.awt.UnoControlNumericFieldModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "NumericField")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("ValueMin", min)
            ctl_props.setPropertyValue("ValueMax", max)
            ctl_props.setPropertyValue("ValueStep", increment)
            ctl_props.setPropertyValue("DecimalAccuracy", accuracy)
            ctl_props.setPropertyValue("Spin", spin_button)
            ctl_props.setPropertyValue("Name", name)
            if not value is None:
                ctl_props.setPropertyValue("Value", value)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlNumericField, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create numeric field control: {e}") from e

    @classmethod
    def insert_progress_bar(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        min: int = 0,
        max: int = 100,
        value: int = 0,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> UnoControlProgressBar:
        """
        Create a new control of type GroupBox in the actual dialog.

        Args:
            dialog_ctrl (XControl): control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height.
            min (float, optional): Specifies the smallest value that can be entered in the control. Defaults to `01``.
            max (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``100``.
            value (int, optional): The value initial value of the progress bar. Defaults to ``0``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create numeric field control control

        Returns:
            UnoControlProgressBar: Group box Control
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlProgressBarModel", msf.createInstance("com.sun.star.awt.UnoControlProgressBarModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "ProgressBar")
            # set properties in the model
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("ProgressValueMin", min)
            ctl_props.setPropertyValue("ProgressValueMax", max)
            ctl_props.setPropertyValue("ProgressValue", value)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlProgressBar, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create numeric field control: {e}") from e

    @classmethod
    def insert_radio_button(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        multiline: bool = False,
        name: str = "",
        **props: Any,
    ) -> UnoControlRadioButton:
        """
        Create a new control of type RadioButton in the actual dialog.

        Args:
            dialog_ctrl (XControl): Control
            label (str): Contents of label
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Default ``20``
            multiline (bool, optional): Specifies if the control can display multiple lines of text. Defaults to ``False``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create radio button

        Returns:
            UnoControlRadioButton: Radio button control
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)

            model = cast(
                "UnoControlRadioButtonModel", msf.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
            )

            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "RadioButton")

            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Label", label)
            ctl_props.setPropertyValue("MultiLine", multiline)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)
            result = cast(UnoControlRadioButton, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create radio button control: {e}") from e

    @classmethod
    def insert_tab_control(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 1,
        name: str = "",
    ) -> UnoControlTabPageContainer:
        """
        Create a new control of type tab in the actual dialog.

        Args:
            dialog_ctrl (XControl): control
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height. Defaults to ``1``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.

        Raises:
            Exception: If unable to create file control control

        Returns:
            UnoControlTabPageContainer: Tab Control
        """
        try:
            dialog = cast(UnoControlDialog, cls.get_dialog(dialog_ctrl))
            if not dialog or not cast(XServiceInfo, dialog).supportsService("com.sun.star.awt.UnoControlDialog"):
                raise Exception("Could not get dialog")
            dialog_model = cast("UnoControlDialogModel", dialog.getModel())
            model = cast(
                "UnoControlTabPageContainerModel",
                dialog_model.createInstance("com.sun.star.awt.tab.UnoControlTabPageContainerModel"),
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog)
            if name:
                model.Name = name
            else:
                model.Name = cls.create_name(name_con, "TabControl")

            # Add the model to the dialog
            dialog_model.insertByName(model.Name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlTabPageContainer, ctrl_con.getControl(model.Name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create Tab control: {e}") from e

    @classmethod
    def insert_tab_page(
        cls,
        dialog_ctrl: XControl,
        *,
        tab_ctrl: UnoControlTabPageContainer,
        title: str,
        tab_position: int,
        name: str = "",
        **props: Any,
    ) -> UnoControlTabPage:
        """
        Create a new control of type Tab in the actual tab control.

        Args:
            tab_ctrl (UnoControlTabPageContainer): Tab Container
            title (str): Tab title
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create file control control

        Returns:
            UnoControlTabPage: Tab Control
        """

        def create_name(ctl: UnoControlTabPageContainer, name: str) -> str:
            items = cast(Tuple[UnoControlTabPage, ...], ctl.Controls)  # type: ignore
            i = 1
            nm = f"{name}{i}"
            for itm in items:
                if not cast(XServiceInfo, itm).supportsService("com.sun.star.awt.tab.UnoControlTabPage"):
                    continue
                itm_model = cast("UnoControlTabPageModel", itm.getModel())
                ctl_props = cls.get_control_props(itm_model)
                prop_name = ctl_props.getPropertyValue("Name")
                if prop_name and prop_name == nm:
                    i += 1
                    nm = f"{name}{i}"
                else:
                    break
            return nm

        try:
            if not tab_ctrl or not cast(XServiceInfo, tab_ctrl).supportsService(
                "com.sun.star.awt.tab.UnoControlTabPageContainer"
            ):
                raise Exception("Not a valid UnoControlTabPageContainer")
            dialog = cast(UnoControlDialog, cls.get_dialog(dialog_ctrl))
            if not dialog or not cast(XServiceInfo, dialog).supportsService("com.sun.star.awt.UnoControlDialog"):
                raise Exception("Could not get dialog")
            dialog_model = cast("UnoControlDialogModel", dialog.getModel())
            model = cast(
                "UnoControlTabPageModel", dialog_model.createInstance("com.sun.star.awt.tab.UnoControlTabPageModel")
            )

            tab_ctrl_model = cast("UnoControlTabPageContainerModel", tab_ctrl.getModel())
            model_init = mLo.Lo.qi(XInitialization, model, True)
            model_init.initialize((tab_position,))

            model.Title = title
            name_con = cls.get_dialog_nm_con(dialog)
            # nm = cls.create_name(name_con, "TabPage")
            if not name:
                name = create_name(tab_ctrl, "TabPage")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            # ctl_props.setPropertyValue("Title", title)
            ctl_props.setPropertyValue("Name", name)
            # ctl_props.setPropertyValue("Border", int(border))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            index = tab_position - 1
            tab_ctrl_model.insertByIndex(index, model)
            controls = cast(Tuple[UnoControlTabPage, ...], tab_ctrl.Controls)  # type: ignore
            if len(controls) > index:
                return controls[index]
            return None  # type: ignore

        except Exception as e:
            raise Exception(f"Could not create Tab control: {e}") from e

    @classmethod
    def insert_text_field(
        cls,
        dialog_ctrl: XControl,
        *,
        text: str,
        x: int,
        y: int,
        width: int,
        height: int = 12,
        echo_char: str = "",
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        **props: Any,
    ) -> UnoControlEdit:
        """
        Inserts a text Field

        Args:
            dialog_ctrl (XControl): Control
            text (str): Text value
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to ``12``.
            echo_char (str, optional): Character used for masking. Must be a single character.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create text field

        Returns:
            UnoControlEdit: Text Field Control

        See Also:
            `API UnoControlEditModel  Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlEditModel.html>`_
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlEditModel", msf.createInstance("com.sun.star.awt.UnoControlEditModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "TextField")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Text", text)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            if len(echo_char) == 1:  # for password fields
                ctl_props.setPropertyValue("EchoChar", ord(echo_char))

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast(UnoControlEdit, ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return result
        except Exception as e:
            raise Exception(f"Could not create text field control: {e}") from e

    # endregion    add components to a dialog

    @staticmethod
    def _set_size_pos(ctl: XWindow, x: int = -1, y: int = -1, width: int = -1, height: int = -1) -> None:
        """
        Set Position and size for a control

        Args:
            ctl (XWindow): Control that implements XWindow
            x (int, optional): X Position. Defaults to -1.
            y (int, optional): Y Position. Defaults to -1.
            width (int, optional): Width. Defaults to -1.
            height (int, optional): Height. Defaults to -1.
        """
        if x < 0 and y < 0 and width < 0 and height < 0:
            return

        pos_size = None
        if x > -1 and y > -1 and width > -1 and height > -1:
            pos_size = PosSize.POSSIZE
        elif x > -1 and y > -1:
            pos_size = PosSize.POS
        elif width > -1 and height > -1:
            pos_size = PosSize.SIZE
        if pos_size is not None:
            ctl.setPosSize(x, y, width, height, pos_size)
