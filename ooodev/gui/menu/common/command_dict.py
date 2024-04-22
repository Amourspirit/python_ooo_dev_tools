from typing import TypedDict


class CommandDict(TypedDict, total=False):
    name: str
    library: str
    language: str
    location: str
    module: str
