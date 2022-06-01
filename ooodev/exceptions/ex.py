# coding: utf-8
from typing import Any

class MissingInterfaceError(Exception):
    def __init__(self, interface: Any, message:Any = None) -> None:
        if message is None:
            try:
                message = f"Missing interface {interface.__pyunointerface__}"
            except AttributeError:
                message = "Missing Uno Interface Error"
        super().__init__(message, interface)
    
    def __str__(self) -> str:
        return repr(self.args[0])


class CellError(Exception):
    pass
