from __future__ import annotations
from typing import Any, Tuple, cast, TYPE_CHECKING
import uno

from ooo.dyn.form.list_source_type import ListSourceType

from ooodev.adapter.form.data_aware_control_model_partial import DataAwareControlModelPartial
from ooodev.adapter.form.update_events import UpdateEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import props as mProps
from ooodev.utils.kind.form_component_kind import FormComponentKind

from ooodev.form.controls.form_ctl_list_box import FormCtlListBox

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form.component import DatabaseListBox as ControlModel  # service
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlDbListBox(FormCtlListBox, DataAwareControlModelPartial, UpdateEvents):
    """``com.sun.star.form.component.DatabaseListBox`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.DatabaseListBox`` service.
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
        FormCtlListBox.__init__(self, ctl=ctl, lo_inst=lo_inst)
        generic_args = self._get_generic_args()
        UpdateEvents.__init__(self, trigger_args=generic_args, cb=self._on_update_events_add_remove)
        DataAwareControlModelPartial.__init__(self, self.get_model())

    # region Lazy Listeners
    def _on_update_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.get_model().addUpdateListener(self.events_listener_update)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    if TYPE_CHECKING:
        # override the methods to provide type hinting
        def get_model(self) -> ControlModel:
            """Gets the model for this control"""
            return cast("ControlModel", super().get_model())

    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        return FormComponentKind.DATABASE_LIST_BOX

    # endregion Overrides

    # region Properties
    if TYPE_CHECKING:
        # override the properties to provide type hinting
        @property
        def model(self) -> ControlModel:
            """Gets the model for this control"""
            return self.get_model()

    @property
    def bound_column(self) -> int:
        """Gets/Sets the bound column.

        Specifies which column of the list result set should be used for data exchange.

        When you make a selection from a list box, the "BoundColumn" property reflects which column value of a result set
        should be used as the value of the component. If the control is bound to a database field, the column value is stored
        in the database field identified by the property ``com.sun.star.form.DataAwareControlModel.DataField``.

        Returns:
            int: If ``-1`` then the index (starting at 0) of the selected list box entry is stored in the current database field.
            If ``0`` or greater, the column value of the result set at the position (0-indexed) is stored in the current database field.
            In particular, for value 0, the selected (displayed) list box string is stored.

        Note:
            The bound column property is only used if a list source is defined and the list source matches with the types
            ``com.sun.star.form.ListSourceType.TABLE``, ``com.sun.star.form.ListSourceType.QUERY``, ``com.sun.star.form.ListSourceType.SQL``
            or ``com.sun.star.form.ListSourceType.SQLPASSTHROUGH``.
            Otherwise the property is ignored, as there is no result set from which to get the column values.
        """
        return self.model.BoundColumn

    @bound_column.setter
    def bound_column(self, value: int) -> None:
        self.model.BoundColumn = value

    @property
    def list_source_type(self) -> ListSourceType:
        """Gets/Sets the list source type

        Returns:
            ListSourceType: The list source type
        """
        return ListSourceType(self.model.ListSourceType)

    @list_source_type.setter
    def list_source_type(self, value: ListSourceType) -> None:
        self.model.ListSourceType = value  # type: ignore

    @property
    def list_source(self) -> Tuple[str, ...]:
        """Gets/Sets the list source

        Returns:
            Tuple[str, ...]: The list source
        """
        return self.model.ListSource

    @list_source.setter
    def list_source(self, value: Tuple[str, ...]) -> None:
        props = self.get_property_set()
        uno.invoke(props, "setPropertyValue", ("ListSource", mProps.Props.any(*value)))  # type: ignore

    @property
    def selected_value(self) -> Any:
        """Gets The selected value, if there is at most one.

        Returns:
            Any: The selected value
        """
        return self.model.SelectedValue

    # endregion Properties
