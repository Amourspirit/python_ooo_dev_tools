from __future__ import annotations
from typing import TYPE_CHECKING, cast, Any
from ooodev.utils.props import Props
from ooodev.utils.gen_util import Util as GenUtil

from com.sun.star.awt import XControl
from com.sun.star.awt import XControlModel

from ooo.dyn.view.selection_type import SelectionType

_CTL_BUTTON = "Button"
_CTL_CHECKBOX = "CheckBox"
_CTL_COMBOBOX = "ComboBox"
_CTL_CURRENCYFIELD = "CurrencyField"
_CTL_DATEFIELD = "DateField"
_CTL_FILECONTROL = "FileControl"
_CTL_FIXED_LINE = "FixedLine"
_CTL_FIXEDTEXT = "FixedText"
_CTL_FORMATTED_FIELD = "FormattedField"
_CTL_GROUPBOX = "GroupBox"
_CTL_HYPERLINK = "Hyperlink"
_CTL_IMAGECONTROL = "ImageControl"
_CTL_LISTBOX = "ListBox"
_CTL_NUMERICFIELD = "NumericField"
_CTL_PATTERNFIELD = "PatternField"
_CTL_PROGRESSBAR = "ProgressBar"
_CTL_RADIOBUTTON = "RadioButton"
_CTL_SCROLLBAR = "ScrollBar"
_CTL_TABLE_CONTROL = "TableControl"
_CTL_TEXTFIELD = "TextField"
_CTL_TIMEFIELD = "TimeField"
_CTL_TREE_CONTROL = "TreeControl"

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButtonModel  # service
    from com.sun.star.awt.tree import TreeControl  # service
    from com.sun.star.awt.tree import TreeControlModel  # service


class DialogControl:
    """Dialog control base class"""

    def __init__(self, ctl: XControl) -> None:
        self._control = ctl
        self._control_model = ctl.getModel()

    # region property_get()
    def _property_get_border(self, control_type: str) -> Any:
        """Get control border property"""
        kinds = {
            _CTL_COMBOBOX,
            _CTL_CURRENCYFIELD,
            _CTL_DATEFIELD,
            _CTL_FILECONTROL,
            _CTL_FIXEDTEXT,
            _CTL_FORMATTED_FIELD,
            _CTL_HYPERLINK,
            _CTL_IMAGECONTROL,
            _CTL_LISTBOX,
            _CTL_NUMERICFIELD,
            _CTL_PATTERNFIELD,
            _CTL_PROGRESSBAR,
            _CTL_SCROLLBAR,
            _CTL_TABLE_CONTROL,
            _CTL_TEXTFIELD,
            _CTL_TIMEFIELD,
            _CTL_TREE_CONTROL,
        }
        if control_type in kinds:
            return Props.get(self._control_model, "Border")

        kinds = {_CTL_CHECKBOX, _CTL_RADIOBUTTON}
        if control_type in kinds:
            return Props.get(self._control_model, "VisualEffect")
        raise ValueError(f"Control border type {control_type} not supported")

    def _property_get_cancel(self, control_type: str) -> Any:
        """Get control cancel property"""
        if control_type == _CTL_BUTTON:
            return Props.get(self._control_model, "PushButtonType")
        raise ValueError(f"Control cancel type {control_type} not supported")

    def _property_get_caption(self, control_type: str) -> Any:
        """Get control caption property"""
        kinds = {
            _CTL_BUTTON,
            _CTL_CHECKBOX,
            _CTL_FIXED_LINE,
            _CTL_FIXEDTEXT,
            _CTL_GROUPBOX,
            _CTL_HYPERLINK,
            _CTL_RADIOBUTTON,
        }
        if control_type in kinds:
            return Props.get(self._control_model, "Label")
        raise ValueError(f"Control caption type {control_type} not supported")

    def _property_get_current_node(self, control_type: str) -> Any:
        """Get control current node"""
        if control_type == _CTL_TREE_CONTROL:
            model = cast("TreeControlModel", self._control_model)
            view = cast("TreeControl", self._control)
            if model.SelectionType == SelectionType.NONE:
                selection = view.getSelection()
                if GenUtil.is_iterable(selection):
                    return selection[0]
                return selection
        raise ValueError(f"Control current node {control_type} not supported")

    def _property_get_default(self, control_type: str) -> Any:
        """Get control default button property"""
        if control_type == _CTL_BUTTON:
            model = cast("UnoControlButtonModel", self._control_model)
            return model.DefaultButton
        raise ValueError(f"Control default button type {control_type} not supported")

    def _property_get_enabled(self) -> Any:
        """Get control cancel property"""
        return Props.get(self._control_model, "Enabled")

    # endregion property_get()
