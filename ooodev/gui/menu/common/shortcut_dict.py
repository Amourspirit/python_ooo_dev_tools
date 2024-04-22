from typing import TypedDict


class ShortcutDict(TypedDict, total=False):
    key: str
    save: bool
