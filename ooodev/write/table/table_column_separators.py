from __future__ import annotations

from typing import Any, cast, TYPE_CHECKING, Tuple
from ooodev.adapter.text.table_column_separator_struct_comp import TableColumnSeparatorStructComp
from ooodev.events.events import Events
from ooodev.utils import gen_util as mGenUtil

if TYPE_CHECKING:
    from com.sun.star.text import TableColumnSeparator  # struct
    from ooodev.events.args.key_val_args import KeyValArgs


class TableColumnSeparators:
    def __init__(self, component: Any) -> None:
        self._component = component
        self._separators = cast(Tuple["TableColumnSeparator", ...], self._component.TableColumnSeparators)
        self._event_provider = Events(self)
        self._props = {}

        def on_comp_struct_changed(src: TableColumnSeparatorStructComp, event_args: KeyValArgs) -> None:
            index = int(event_args.event_data["prop_name"])
            lst = list(self._separators)
            self._props[index] = src
            lst[index] = src.copy()
            self._component.TableColumnSeparators = tuple(lst)
            self._separators = self._component.TableColumnSeparators

        self._fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member
        self._event_provider.subscribe_event(
            "com_sun_star_text_TableColumnSeparator_changed", self._fn_on_comp_struct_changed
        )

    def __iter__(self):
        """Iterates through the separators."""
        for i in range(len(self._separators)):
            yield self[i]

    def __len__(self) -> int:
        """Gets the number of separators."""
        return len(self._separators)

    def __getitem__(self, key: int) -> TableColumnSeparatorStructComp:
        """Gets the separator at the specified index."""
        index = self._get_index(key, True)
        if index in self._props:
            return self._props[index]
        struct = self._separators[index]
        comp = TableColumnSeparatorStructComp(
            component=struct, prop_name=f"{index}", event_provider=self._event_provider
        )
        self._props[index] = comp
        return comp

    def _get_index(self, idx: int, allow_greater: bool = False) -> int:
        """
        Gets the index.

        Args:
            idx (int): Index of sheet. Can be a negative value to index from the end of the list.
            allow_greater (bool, optional): If True and index is greater then the number of
                sheets then the index becomes the next index if sheet were appended. Defaults to False.

        Returns:
            int: Index value.
        """
        count = len(self._separators)
        return mGenUtil.Util.get_index(idx, count, allow_greater)
