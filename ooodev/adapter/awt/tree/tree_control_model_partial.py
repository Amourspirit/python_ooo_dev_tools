from __future__ import annotations
from typing import TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooo.dyn.view.selection_type import SelectionType
from ooodev.units.unit_px import UnitPX
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.adapter.awt.uno_control_model_partial import UnoControlModelPartial

if TYPE_CHECKING:
    from com.sun.star.awt.tree import TreeControlModel  # Service
    from com.sun.star.awt.tree import XTreeDataModel
    from ooodev.units.unit_obj import UnitT


class TreeControlModelPartial(UnoControlModelPartial):
    """Partial class for TreeControlModel."""

    def __init__(self) -> None:
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.TreeControlModel`` service.
        """
        # pylint: disable=unused-argument
        if not isinstance(self, ModelPropPartial):
            raise TypeError("This class must be used as a mixin that implements ModelPropPartial.")

        self.model: TreeControlModel
        UnoControlModelPartial.__init__(self)

    # region Properties

    @property
    def data_model(self) -> XTreeDataModel:
        """
        Specifies the ``XTreeDataModel`` that is providing the hierarchical data.

        You can implement your own instance of ``XTreeDataModel`` or use the ``MutableTreeDataModel``.
        """
        return self.model.DataModel

    @data_model.setter
    def data_model(self, value: XTreeDataModel) -> None:
        self.model.DataModel = value

    @property
    def editable(self) -> bool:
        """
        Gets/Sets whether the nodes of the tree are editable.

        The default value is ``False``.
        """
        return self.model.Editable

    @editable.setter
    def editable(self, value: bool) -> None:
        self.model.Editable = value

    @property
    def invokes_stop_node_editing(self) -> bool:
        """
        Gets/Sets what happens when editing is interrupted by selecting another node in the tree,
        a change in the tree's data, or by some other means.

        Setting this property to ``True`` causes the changes to be automatically saved when editing is interrupted.
        ``False`` means that editing is canceled and changes are lost

        The default value is ``False``.
        """
        return self.model.InvokesStopNodeEditing

    @invokes_stop_node_editing.setter
    def invokes_stop_node_editing(self, value: bool) -> None:
        self.model.InvokesStopNodeEditing = value

    @property
    def root_displayed(self) -> bool:
        """
        Gets/Sets if the root node of the tree is displayed.

        If ``RootDisplayed`` is set to ``False``, the root node of a model is no longer a valid node for the
        ``XTreeControl`` and can't be used with any method of ``XTreeControl``.

        The default value is ``True``.
        """
        return self.model.RootDisplayed

    @root_displayed.setter
    def root_displayed(self, value: bool) -> None:
        self.model.RootDisplayed = value

    @property
    def row_height(self) -> UnitPX:
        """
        Gets/Sets the height of each row, in pixels units.

        If the specified value is less than or equal to zero, the row height is the maximum height of all rows.

        The default value is ``0``

        Returns:
            UnitPX: Row height in pixels.

        Note:
            Value can be set as an integer or a ``UnitPX`` instance.
        """
        return UnitPX(self.model.RowHeight)

    @row_height.setter
    def row_height(self, value: int | UnitT) -> None:
        px = UnitPX.from_unit_val(value)
        self.model.RowHeight = int(px.value)

    @property
    def selection_type(self) -> SelectionType:
        """
        Gets/Sets the selection mode that is enabled for this tree.

        The default value is ``com.sun.star.view.SelectionType.NONE``

        Hint:
            - ``SelectionType`` can be imported from ``ooo.dyn.view.selection_type``
        """
        return self.model.SelectionType  # type: ignore

    @selection_type.setter
    def selection_type(self, value: SelectionType) -> None:
        self.model.SelectionType = value  # type: ignore

    @property
    def shows_handles(self) -> bool:
        """
        Gets/Sets whether the node handles should be displayed.
        The handles are doted lines that visualize the tree like hierarchy.

        The default value is ``True``.
        """
        return self.model.ShowsHandles

    @shows_handles.setter
    def shows_handles(self, value: bool) -> None:
        self.model.ShowsHandles = value

    @property
    def shows_root_handles(self) -> bool:
        """
        Gets/Sets whether the node handles should also be displayed at root level.

        The default value is ``True``.
        """
        return self.model.ShowsRootHandles

    @shows_root_handles.setter
    def shows_root_handles(self, value: bool) -> None:
        self.model.ShowsRootHandles = value

    # endregion Properties
