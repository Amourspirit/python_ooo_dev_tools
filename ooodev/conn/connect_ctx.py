from __future__ import annotations
from typing import cast

from com.sun.star.uno import XComponentContext
from ooodev.utils.typing.over import override
from ooodev.conn.connect import ConnectBase


class ConnectCtx(ConnectBase):
    """
    Connection to LibreOffice/OpenOffice

    This class is used to connect to LibreOffice/OpenOffice.
    It is used to create a connection to LibreOffice/OpenOffice.
    It is used to create a connection to LibreOffice/OpenOffice.

    .. versionadded:: 0.53.0
    """

    @override
    def __init__(self, ctx: XComponentContext) -> None:
        super().__init__()
        self._ctx = cast(XComponentContext, ctx)

    @override
    def connect(self) -> None:
        self.log.info("connect() Connection Established")

    @override
    def kill_soffice(self) -> None:
        raise NotImplementedError("kill_soffice is not implemented in this child class")
