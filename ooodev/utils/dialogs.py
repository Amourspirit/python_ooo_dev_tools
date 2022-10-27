# coding: utf-8
# region Imports
from __future__ import annotations
from typing import TYPE_CHECKING, Any, Iterable, Tuple, cast
from enum import IntEnum
from . import lo as mLo

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
from com.sun.star.lang import XMultiServiceFactory

# com.sun.star.awt.PushButtonType
from ooo.dyn.awt.push_button_type import PushButtonType

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButtonModel  # service
    from com.sun.star.awt import UnoControlCheckBoxModel  # service
    from com.sun.star.awt import UnoControlComboBoxModel  # service
    from com.sun.star.awt import UnoControlEditModel  # service
    from com.sun.star.awt import UnoControlFixedTextModel  # service
    from com.sun.star.container import XNameAccess
    from com.sun.star.lang import EventObject
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
            print(f"  Defalut Contol: {cls.get_control_class_id(ctl)}")
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
    def create_dialog_control(
        cls,
        x: int,
        y: int,
        width: int,
        height: int,
        title: str,
        **props: Any,
    ) -> XControl:
        """
        Creates a dialog control

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
            XControl: Control

        See Also:
            `API UnoControlDialogModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogModel.html>`_
        """
        try:
            dialog_ctrl = mLo.Lo.create_instance_mcf(XControl, "com.sun.star.awt.UnoControlDialog", raise_err=True)
            xcontrol_model = mLo.Lo.create_instance_mcf(
                XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True
            )
            dialog_ctrl.setModel(xcontrol_model)

            cprops = cls.get_control_props(dialog_ctrl.getModel())

            cprops.setPropertyValue("PositionX", x)
            cprops.setPropertyValue("PositionY", y)
            cprops.setPropertyValue("Height", height)
            cprops.setPropertyValue("Width", width)

            cprops.setPropertyValue("Title", title)
            cprops.setPropertyValue("Name", "OfficeDialog")

            cprops.setPropertyValue("Step", 0)
            cprops.setPropertyValue("Moveable", True)
            cprops.setPropertyValue("TabIndex", 0)

            # set any extra user properties
            for k, v in props.items():
                cprops.setPropertyValue(k, v)

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
        dialog = cls.get_dialog(dialog_ctrl)
        # dialog_component.dispose() # free window resources
        # commented out or the Add-on dialog crashes when called a second time
        # because createPeer() cannot find a model

        return dialog

    # endregion create a dialog

    # region    add components to a dialog

    @classmethod
    def insert_label(
        cls,
        dialog_ctrl: XControl,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 8,
        **props: Any,
    ) -> XControl:
        """
        Insert a label into a control

        Args:
            dialog_ctrl (XControl): Control
            label (str): Contents of label
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Default 8
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create label

        Returns:
            XControl: control

        See Also:
            `API UnoControlFixedTextModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlFixedTextModel.html>`_
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)

            model = cast("UnoControlFixedTextModel", msf.createInstance("com.sun.star.awt.UnoControlFixedTextModel"))

            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            nm = cls.create_name(name_con, "FixedText")

            cprops = cls.get_control_props(model)
            cprops.setPropertyValue("PositionX", x)
            cprops.setPropertyValue("PositionY", y + 2)
            cprops.setPropertyValue("Height", height)
            cprops.setPropertyValue("Width", width)
            cprops.setPropertyValue("Label", label)
            cprops.setPropertyValue("Name", nm)

            # set any extra user properties
            for k, v in props.items():
                cprops.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(nm, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl)
            return ctrl_con.getControl(nm)
        except Exception as e:
            raise Exception(f"Could not create fixed text control: {e}") from e

    @staticmethod
    def get_dialog_nm_con(dialog_ctrl: XControl) -> XNameContainer:
        """
        Gets Name container from control

        Args:
            dialog_ctrl (XControl): Dialog control

        Returns:
            XNameContainer: Name Container
        """
        return mLo.Lo.qi(XNameContainer, dialog_ctrl.getModel(), True)

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
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 14,
        btn_type: PushButtonType | None = None,
        **props: Any,
    ) -> XControl:
        """
        Insert Button Control

        Args:
            dialog_ctrl (XControl): control
            label (str): Button Label
            x (int): X coordinate
            y (int): Y coordinate
            width (int): width
            height (int, optional): Height. Defaults to 14.
            btn_type (PushButtonType | None, optional): Type of Button.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create button control

        Returns:
            XControl: Button control

        See Also:
            `API UnoControlButtonModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlButtonModel.html>`_
        """
        if btn_type is None:
            btn_type = PushButtonType.STANDARD
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlButtonModel", msf.createInstance("com.sun.star.awt.UnoControlButtonModel"))

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            nm = cls.create_name(name_con, "CommandButton")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            cprops = cls.get_control_props(model)

            cprops.setPropertyValue("PositionX", x)
            cprops.setPropertyValue("PositionY", y)
            cprops.setPropertyValue("Height", height)
            cprops.setPropertyValue("Width", width)
            cprops.setPropertyValue("Label", label)
            # cprops.setPropertyValue("PushButtonType", btn_type)
            uany = uno.Any("short", btn_type)
            uno.invoke(cprops, "setPropertyValue", ("PushButtonType", uany))
            cprops.setPropertyValue("Name", nm)

            # set any extra user properties
            for k, v in props.items():
                cprops.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(nm, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl)

            # use the model's name to get its view inside the dialog
            return ctrl_con.getControl(nm)
        except Exception as e:
            raise Exception(f"Could not create button control: {e}") from e

    @classmethod
    def insert_text_field(
        cls,
        dialog_ctrl: XControl,
        text: str,
        x: int,
        y: int,
        width: int,
        height=12,
        echo_char: str = "",
        **props: Any,
    ) -> XControl:
        """
        Inserts a text Field

        Args:
            dialog_ctrl (XControl): Control
            text (str): Text value
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to 12.
            echo_char (str, optional): Character used for masking. Must be a single character.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create text field

        Returns:
            XControl: Text Field Control

        See Also:
            `API UnoControlEditModel  Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlEditModel.html>`_
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlEditModel", msf.createInstance("com.sun.star.awt.UnoControlEditModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            nm = cls.create_name(name_con, "TextField")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            cprops = cls.get_control_props(model)

            cprops.setPropertyValue("PositionX", x)
            cprops.setPropertyValue("PositionY", y)
            cprops.setPropertyValue("Height", height)
            cprops.setPropertyValue("Width", width)
            cprops.setPropertyValue("Text", text)
            cprops.setPropertyValue("Name", nm)

            # set any extra user properties
            for k, v in props.items():
                cprops.setPropertyValue(k, v)

            if len(echo_char) == 1:  # for password fields
                cprops.setPropertyValue("EchoChar", ord(echo_char))

            # Add the model to the dialog
            name_con.insertByName(nm, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl)

            # use the model's name to get its view inside the dialog
            return ctrl_con.getControl(nm)
        except Exception as e:
            raise Exception(f"Could not create text field control: {e}") from e

    @classmethod
    def insert_password_field(
        cls, dialog_ctrl: XControl, text: str, x: int, y: int, width: int, height=12, **props: Any
    ) -> XControl:
        """
        Inserts a password field.

        Args:
            dialog_ctrl (XControl): Control
            text (str): Text value
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to 12.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create text field

        Returns:
            XControl: Text Field Control

        See Also:
            :py:meth:`~.dialogs.Dialogs.insert_text_field`
        """
        return cls.insert_text_field(
            dialog_ctrl=dialog_ctrl, text=text, x=x, y=y, width=width, height=height, echo_char="*", **props
        )

    @classmethod
    def insert_combo_box(
        cls,
        dialog_ctrl: XControl,
        entries: Iterable[str],
        x: int,
        y: int,
        width: int,
        height: int = 12,
        max_text_len: int = 0,
        drop_down: bool = True,
        read_only: bool = False,
        **props: Any,
    ) -> XControl:
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
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create combo box control

        Returns:
            XControl: Combo box control

        See Also:
            `API UnoControlComboBoxModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlComboBoxModel.html>`_
        """
        try:
            if max_text_len < 0:
                max_text_len = 0
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlComboBoxModel", msf.createInstance("com.sun.star.awt.UnoControlComboBoxModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            nm = cls.create_name(name_con, "ComboBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            cprops = cls.get_control_props(model)

            cprops.setPropertyValue("PositionX", x)
            cprops.setPropertyValue("PositionY", y)
            cprops.setPropertyValue("Height", height)
            cprops.setPropertyValue("Width", width)
            cprops.setPropertyValue("Name", nm)
            cprops.setPropertyValue("Dropdown", drop_down)
            cprops.setPropertyValue("StringItemList", entries)
            cprops.setPropertyValue("MaxTextLen", max_text_len)
            cprops.setPropertyValue("ReadOnly", read_only)

            # set any extra user properties
            for k, v in props.items():
                cprops.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(nm, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl)

            # use the model's name to get its view inside the dialog
            return ctrl_con.getControl(nm)
        except Exception as e:
            raise Exception(f"Could not create combo box control: {e}") from e

    @classmethod
    def insert_check_box(
        cls,
        dialog_ctrl: XControl,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 8,
        tri_state: bool = True,
        state: Dialogs.StateEnum = StateEnum.CHECKED,
        **props: Any,
    ) -> XControl:
        """
        Inserts a check box control

        Args:
            dialog_ctrl (XControl): Control
            label (str): Checkbox label text
            x (int): X coordinate
            y (int): Y coordinate
            width (int): Width
            height (int, optional): Height. Defaults to 8.
            tri_state (StateEnum, optional): Specifies that the control may have the state "don't know". Defaults to ``StateEnum.CHECKED``.
            state (int, optional): specifies the state of the control. Defaults to 1.
            props (dict, optional): Extra properties to set for control.

        Raises:
            Exception: If unable to create checkbox control.

        Returns:
            XControl: Check box control

        See Also:
            `API UnoControlCheckBoxModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlCheckBoxModel.html>`_
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlCheckBoxModel", msf.createInstance("com.sun.star.awt.UnoControlCheckBoxModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            nm = cls.create_name(name_con, "CheckBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            cprops = cls.get_control_props(model)

            cprops.setPropertyValue("PositionX", x)
            cprops.setPropertyValue("PositionY", y)
            cprops.setPropertyValue("Height", height)
            cprops.setPropertyValue("Width", width)
            cprops.setPropertyValue("Name", nm)
            cprops.setPropertyValue("Label", label)
            cprops.setPropertyValue("TriState", tri_state)
            cprops.setPropertyValue("State", int(state))

            # set any extra user properties
            for k, v in props.items():
                cprops.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(nm, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl)

            # use the model's name to get its view inside the dialog
            return ctrl_con.getControl(nm)
        except Exception as e:
            raise Exception(f"Could not create check box control: {e}") from e

    # endregion    add components to a dialog
