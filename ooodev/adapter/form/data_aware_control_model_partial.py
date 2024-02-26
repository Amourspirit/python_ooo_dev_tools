from __future__ import annotations

from typing import cast, TYPE_CHECKING

if TYPE_CHECKING:
    from com.sun.star.form import DataAwareControlModel  # service
    from com.sun.star.beans import XPropertySet


class DataAwareControlModelPartial:
    """
    Class is an abstract service for specialized FormControlModels which are data aware and thus can be bound to a data source
    """

    # pylint: disable=unused-argument

    def __init__(self, component: DataAwareControlModel) -> None:
        """
        Constructor

        Args:
            component (SheetCellCursor): UNO Sheet Cell Cursor Component
        """
        self.__component = component

    # region Properties
    @property
    def bound_field(self) -> XPropertySet:
        """Gets/Sets the name of the field in the data source to which the control is bound."""
        return self.__component.BoundField

    @bound_field.setter
    def bound_field(self, value: XPropertySet) -> None:
        self.__component.BoundField = value

    @property
    def label_control(self) -> XPropertySet:
        """Gets/Sets references to a control model within the same document which should be used as a label."""
        return self.__component.LabelControl

    @label_control.setter
    def label_control(self, value: XPropertySet) -> None:
        self.__component.LabelControl = value

    @property
    def data_field(self) -> str:
        """Gets/Sets the name of the field in the data source to which the control is bound."""
        return cast(str, self.__component.DataField)

    @data_field.setter
    def data_field(self, value: str) -> None:
        self.__component.DataField = value

    @property
    def input_required(self) -> bool:
        """Gets/Sets whether or not input into this field is required, when it is actually bound to a database field."""
        return cast(bool, self.__component.InputRequired)

    @input_required.setter
    def input_required(self, value: bool) -> None:
        self.__component.InputRequired = value

    # endregion Properties
