from __future__ import annotations
from ooodev.loader.inst.options import Options
from ooodev.conn import cache as mCache
from ooodev.conn import connectors
from ooodev.conn.connect import ConnectBase
from ooodev.conn.connect import LoPipeStart
from ooodev.conn.connect import LoSocketStart
from ooodev.conn.connect import LoDirectStart
from ooodev.exceptions import ex as mEx


class LoLoader:
    """Makes a connection to Office"""

    def __init__(
        self,
        connector: connectors.ConnectPipe | connectors.ConnectSocket | ConnectBase | None = None,
        cache_obj: mCache.Cache | None = None,
        opt: Options | None = None,
    ) -> None:
        """
        Loads Office

        Not available in a macro.

        If running outside of office then a bridge is created that connects to office.

        If running from inside of office e.g. in a macro, then ``Lo.XSCRIPTCONTEXT`` is used.
        ``using_pipes`` is ignored with running inside office.

        Args:
            connector (connectors.ConnectPipe, connectors.ConnectSocket, ConnectBase, None): Connection information. Ignore for macros.
            cache_obj (Cache, optional): Cache instance that determines of LibreOffice profile is to be copied and cached
                Ignore for macros. Defaults to None.
            opt (Options, optional): Extra Load options.
        """
        self._lo_inst: ConnectBase | None = None
        if opt is None:
            self._opt = Options()
        else:
            self._opt = opt

        b_connector = connector
        if b_connector is None:
            try:
                self._lo_inst = LoDirectStart()
                self._lo_inst.connect()
            except Exception as e:
                raise mEx.ConnectionError(
                    "Office context could not be created. A connector must be supplied if not running as a macro"
                ) from e
        elif isinstance(b_connector, connectors.ConnectPipe):
            try:
                self._lo_inst = LoPipeStart(connector=b_connector, cache_obj=cache_obj)
                self._lo_inst.connect()
            except Exception as e:
                raise mEx.ConnectionError("Office context could not be created") from e
        elif isinstance(b_connector, connectors.ConnectSocket):
            try:
                self._lo_inst = LoSocketStart(connector=b_connector, cache_obj=cache_obj)
                self._lo_inst.connect()
            except Exception as e:
                raise mEx.ConnectionError("Office context could not be created") from e
        elif isinstance(b_connector, ConnectBase):
            self._lo_inst = b_connector
            if not self._lo_inst.has_connection:
                try:
                    self._lo_inst.connect()
                except Exception as e:
                    raise mEx.ConnectionError("Office context could not be created") from e
        else:
            raise mEx.NotSupportedError("Invalid Connector type. Fatal Error.")

    @property
    def lo_inst(self) -> ConnectBase:
        if self._lo_inst is None:
            raise mEx.LoadingError("Office is not loaded")
        return self._lo_inst

    @property
    def options(self) -> Options:
        return self._opt
