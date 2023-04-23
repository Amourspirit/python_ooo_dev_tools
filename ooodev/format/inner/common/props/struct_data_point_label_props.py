from __future__ import annotations
from typing import NamedTuple
from .prop_pair import PropPair as PropPair


class StructDataPointLabelProps(NamedTuple):
    """Internal Properties"""

    show_number: str
    show_number_in_percent: str
    show_category_name: str
    show_legend_symbol: str
    show_custom_label: str
    show_series_name: str
