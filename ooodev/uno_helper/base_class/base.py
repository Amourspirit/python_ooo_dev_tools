from __future__ import annotations
from typing import Any
import uno
import unohelper
from com.sun.star.uno import XInterface


class Base(unohelper.Base, XInterface):

    # region XInterface
    def acquire(self) -> None:
        pass

    def release(self) -> None:
        pass

    def queryInterface(self, a_type: Any) -> Any:
        if a_type in self.getTypes():
            return self
        return None

    # end region XInterface
