from __future__ import annotations
from typing import cast, Any, TYPE_CHECKING
import uno
from com.sun.star.lang import XServiceInfo
from ooo.dyn.form.form_component_type import FormComponentType
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.awt import XControlModel
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.drawing import XControlShape
    from com.sun.star.drawing import ControlShape  # service

# pylint: disable=import-outside-toplevel


class FormControlFactory(LoInstPropsPartial):

    def __init__(self, draw_page: XDrawPage, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        self._draw_page = draw_page
        self._doc = mLo.Lo.current_doc

    def find_shape_for_control(self, ctl_model: XControlModel) -> XControlShape | None:
        """
        Find the shape for a control.

        Args:
            ctl_model (XControlModel): control to find shape for.

        Returns:
            XControlShape | None: Shape for the control or ``None`` if not found.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.adapter.container.index_access_comp import IndexAccessComp

        ia = cast(IndexAccessComp["XControlShape"], IndexAccessComp(self._draw_page))  # type: ignore
        for shape in ia:
            if shape.supportsService("com.sun.star.drawing.ControlShape"):  # type: ignore
                cs = cast("ControlShape", shape)
                if ctl_model == cs.getControl():
                    return cs
        return None

    def _get_id(self, ctl_model: XControlModel) -> int:
        """
        Gets class id for a form component.

        |lo_safe|

        Args:
            ctl_model (XControlModel | FormCtlBase): Control Model.

        Returns:
            int: Class Id if found, Otherwise ``-1``.
        """
        class_id = mProps.Props.get(ctl_model, "ClassId", None)
        return -1 if class_id is None else int(class_id)

    def _get_control(self, ctl_model: XControlModel) -> XControl | None:
        """
        Gets the control from the model.

        Args:
            ctl_model (XControlModel): Control Model.

        Returns:
            XControl | None: Control if found, Otherwise ``None``.
        """
        # pylint: disable=no-member
        control_access = self._doc.get_control_access()
        return None if control_access is None else control_access.getControl(ctl_model)

    def get_control_from_model(self, ctl_model: XControlModel) -> Any:
        """
        Gets the control from the model.

        Args:
            ctl_model (XControlModel): Control Model.

        Returns:
            Any: Control if found, Otherwise ``None``.
        """
        control = self._get_control(ctl_model)
        return None if control is None else self.get_control(control)

    def get_control(self, ctl: XControl) -> Any:
        if ctl is None:
            return None
        model = ctl.getModel()

        control_id = self._get_id(model)
        if control_id == -1:
            return None
        service_info = mLo.Lo.qi(XServiceInfo, ctl)
        if control_id == FormComponentType.COMMANDBUTTON:
            from ooodev.form.controls.form_ctl_button import FormCtlButton

            form_ctl = FormCtlButton(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl
        elif control_id == FormComponentType.CHECKBOX:
            from ooodev.form.controls.form_ctl_check_box import FormCtlCheckBox

            form_ctl = FormCtlCheckBox(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.COMBOBOX:
            from ooodev.form.controls.form_ctl_combo_box import FormCtlComboBox

            form_ctl = FormCtlComboBox(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.CURRENCYFIELD:
            from ooodev.form.controls.form_ctl_currency_field import FormCtlCurrencyField

            form_ctl = FormCtlCurrencyField(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.DATEFIELD:
            from ooodev.form.controls.form_ctl_date_field import FormCtlDateField

            form_ctl = FormCtlDateField(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.FILECONTROL:
            from ooodev.form.controls.form_ctl_file import FormCtlFile

            form_ctl = FormCtlFile(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.FIXEDTEXT:
            from ooodev.form.controls.form_ctl_fixed_text import FormCtlFixedText

            form_ctl = FormCtlFixedText(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.TEXTFIELD:
            if service_info is not None and service_info.supportsService("com.sun.star.form.component.FormattedField"):
                from ooodev.form.controls.form_ctl_formatted_field import FormCtlFormattedField

                form_ctl = FormCtlFormattedField(ctl, lo_inst=self.lo_inst)
                form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
                return form_ctl

            from ooodev.form.controls.form_ctl_text_field import FormCtlTextField

            form_ctl = FormCtlTextField(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.GRIDCONTROL:
            from ooodev.form.controls.form_ctl_grid import FormCtlGrid

            form_ctl = FormCtlGrid(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.LISTBOX:
            from ooodev.form.controls.form_ctl_list_box import FormCtlListBox

            form_ctl = FormCtlListBox(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.GROUPBOX:
            from ooodev.form.controls.form_ctl_group_box import FormCtlGroupBox

            form_ctl = FormCtlGroupBox(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.HIDDENCONTROL:
            from ooodev.form.controls.form_ctl_hidden import FormCtlHidden

            form_ctl = FormCtlHidden(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.IMAGEBUTTON:
            from ooodev.form.controls.form_ctl_image_button import FormCtlImageButton

            form_ctl = FormCtlImageButton(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.LISTBOX:
            from ooodev.form.controls.form_ctl_list_box import FormCtlListBox

            form_ctl = FormCtlListBox(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.NAVIGATIONBAR:
            from ooodev.form.controls.form_ctl_navigation_tool_bar import FormCtlNavigationToolBar

            form_ctl = FormCtlNavigationToolBar(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.NUMERICFIELD:
            from ooodev.form.controls.form_ctl_numeric_field import FormCtlNumericField

            form_ctl = FormCtlNumericField(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.PATTERNFIELD:
            from ooodev.form.controls.form_ctl_pattern_field import FormCtlPatternField

            form_ctl = FormCtlPatternField(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.RADIOBUTTON:
            from ooodev.form.controls.form_ctl_radio_button import FormCtlRadioButton

            form_ctl = FormCtlRadioButton(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.SCROLLBAR:
            from ooodev.form.controls.form_ctl_scroll_bar import FormCtlScrollBar

            form_ctl = FormCtlScrollBar(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.SPINBUTTON:
            from ooodev.form.controls.form_ctl_spin_button import FormCtlSpinButton

            form_ctl = FormCtlSpinButton(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        elif control_id == FormComponentType.TIMEFIELD:
            from ooodev.form.controls.form_ctl_time_field import FormCtlTimeField

            form_ctl = FormCtlTimeField(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        if service_info is None:
            return None

        if service_info.supportsService("com.sun.star.form.component.RichTextControl"):
            from ooodev.form.controls.form_ctl_rich_text import FormCtlRichText

            form_ctl = FormCtlRichText(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl
        elif service_info.supportsService("com.sun.star.form.component.SubmitButton"):
            from ooodev.form.controls.form_ctl_submit_button import FormCtlSubmitButton

            form_ctl = FormCtlSubmitButton(ctl, lo_inst=self.lo_inst)
            form_ctl.control_shape = self.find_shape_for_control(model)  # type: ignore
            return form_ctl

        return None
