from typing import NamedTuple


class CmdData(NamedTuple):
    """Represents a uno command data."""

    command: str
    label: str
    name: str
    popup: bool
    properties: int
    popup_label: str
    tooltip_label: str
    target_url: str
    is_experimental: bool
    module_hotkey: str
    global_hotkey: str
