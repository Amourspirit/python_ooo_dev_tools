from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.events import Events
from ooodev.utils import info as mInfo
from ooodev.adapter.table.table_border2_struct_comp import TableBorder2StructComp

if TYPE_CHECKING:
    from com.sun.star.text import TextTable  # service
    from ooo.dyn.table.table_border2 import TableBorder2
    from ooodev.events.args.key_val_args import KeyValArgs


class WriteTablePropertiesPartial:
    """
    Partial Properties class for TextTable Service.

    This class is used along with ``ooodev.adapter.text.text_table_properties_partial.TextTablePropertiesPartial`` class.
    """

    def __init__(self, component: TextTable) -> None:
        """
        Constructor

        Args:
            component (TextTable): UNO Component that implements ``com.sun.star.text.TextTable`` service.
        """
        self.__component = component
        self.__event_provider = Events(self)
        self.__props = {}

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self.__component, prop_name):
                setattr(self.__component, prop_name, event_args.source.component)

        self.__fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member

        self.__event_provider.subscribe_event(
            "com_sun_star_table_TableBorder2_changed", self.__fn_on_comp_struct_changed
        )

    # region Properties
    @property
    def table_border2(self) -> TableBorder2StructComp:
        """
        Gets/Sets a description of the cell or cell range border.

        If used with a cell range, the top, left, right, and bottom lines are at the edges of the entire range, not at the edges of the individual cell.

        Setting value can be done with a ``TableBorder2`` or ``TableBorder2StructComp`` object.

        Returns:
            TableBorder2StructComp: Table Border.

        Hint:
            - ``TableBorder2`` can be imported from ``ooo.dyn.table.table_border2``
        """
        key = "TableBorder2"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = TableBorder2StructComp(
                self.__component.TableBorder2,  # type: ignore
                key,
                self.__event_provider,
            )
            self.__props[key] = prop
        return cast(TableBorder2StructComp, prop)

    @table_border2.setter
    def table_border2(self, value: TableBorder2 | TableBorder2StructComp) -> None:
        key = "TableBorder2"
        if mInfo.Info.is_instance(value, TableBorder2StructComp):
            self.__component.TableBorder2 = value.copy()  # type: ignore
        else:
            self.__component.TableBorder2 = cast("TableBorder2", value)  # type: ignore
        if key in self.__props:
            del self.__props[key]

    # endregion Properties
