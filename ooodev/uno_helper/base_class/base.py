from __future__ import annotations
from typing import Any

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

import unohelper
from com.sun.star.uno import XInterface


class Base(unohelper.Base, XInterface):

    # region XInterface
    @override
    def acquire(self) -> None:
        raise NotImplementedError

    @override
    def release(self) -> None:
        raise NotImplementedError

    @override
    def queryInterface(self, aType: Any) -> Any:
        if aType in self.getTypes():
            return self
        return None

    # end region XInterface
