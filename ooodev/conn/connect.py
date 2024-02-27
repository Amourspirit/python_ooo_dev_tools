"""Connection to LibreOffice/OpenOffice"""

from __future__ import annotations
from typing import Any, List, TYPE_CHECKING, cast
import contextlib
import os
import time
from abc import ABC, abstractmethod
import subprocess
import signal
from pathlib import Path
import uno
from com.sun.star.connection import NoConnectException  # type: ignore
from ooodev.conn import connectors
from ooodev.conn import cache
from ooodev.utils.sys_info import SysInfo


if TYPE_CHECKING:
    from com.sun.star.connection import XConnector
    from com.sun.star.beans import XPropertySet
    from com.sun.star.bridge import XBridgeFactory
    from com.sun.star.bridge import XBridge
    from com.sun.star.lang import XMultiComponentFactory
    from com.sun.star.uno import XComponentContext
    from com.sun.star.lang import XComponent
    from com.sun.star.bridge import UnoUrlResolver  # service
else:
    XConnector = Any
    XBridgeFactory = Any
    XBridge = Any
    XPropertySet = Any
    XMultiComponentFactory = Any
    XComponentContext = Any
    XComponent = Any
    UnoUrlResolver = Any


class ConnectBase(ABC):
    """Base Abstract Class for all connections to LO"""

    def __init__(self):
        # https://tinyurl.com/yb897bxw
        # https://tinyurl.com/ybk7zqcg

        # start openoffice process with python to use with pyuno using subprocess
        # see https://tinyurl.com/y5y66462
        self._ctx = cast(XComponentContext, None)

    def __eq__(self, other: object) -> bool:
        return NotImplemented

    @abstractmethod
    def connect(self):
        """
        Makes a connection to soffice

        Raises:
            NoConnectException: if unable to obtain a connection to soffice
        """
        ...

    def kill_soffice(self) -> None:
        """
        Attempts to kill instance of soffice created by this instance
        """
        raise NotImplementedError("kill_soffice is not implemented in this child class")

    @property
    def ctx(self) -> XComponentContext:
        """Gets instance Component Context"""
        return self._ctx

    @property
    def start_office(self) -> bool:
        """For compatibility. Returns ``True``"""
        return True

    @property
    def no_restore(self) -> bool:
        """For compatibility. Returns ``True``"""
        return True

    @property
    def no_first_start_wizard(self) -> bool:
        """For compatibility. Returns ``True``"""
        return True

    @property
    def no_logo(self) -> bool:
        """For compatibility. Returns ``True``"""
        return True

    @property
    def invisible(self) -> bool:
        """For compatibility. Returns ``False``"""
        return False

    @property
    def headless(self) -> bool:
        """For compatibility. Returns ``False``"""
        return False

    @property
    def start_as_service(self) -> bool:
        """For compatibility. Returns ``False``"""
        return False

    @property
    def has_connection(self) -> bool:
        """Returns ``True`` if a connection to soffice has been established"""
        return self._ctx is not None

    @property
    def is_remote(self) -> bool:
        """Returns False"""
        return False


class LoBridgeCommon(ConnectBase):
    """Base Abstract Class for LoSocketStart and LoPipeStart"""

    def __init__(self, connector: connectors.ConnectorBridgeBase, cache_obj: cache.Cache | None):
        super().__init__()
        self._connector = connector
        self._soffice_process = None
        self._bridge_component = cast(XComponent, None)
        self._platform = SysInfo.get_platform()
        self._environment = os.environ.copy()
        self._timeout = 30.0
        self._conn_try_sleep = 0.2
        self._cache = cache.Cache(use_cache=False) if cache_obj is None else cache_obj
        if self._cache.use_cache:
            self._environment["TMPDIR"] = str(self._cache.working_dir)
        if connector.env_vars:
            self._environment.update(connector.env_vars)

    @abstractmethod
    def _get_connection_str(self) -> str:
        """
        Gets connection string.

        Such as ``uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager``
        """
        ...

    @abstractmethod
    def _get_connection_identifier(self) -> str:
        """
        Gets connection identifier

        Such as ``socket,host=localhost,port=2002``
        """
        ...

    @abstractmethod
    def _get_bridge(self, local_factory: XMultiComponentFactory, local_ctx: XComponentContext) -> XBridge: ...

    def _connect(self):
        # see also: _connect_alternative()
        conn_str = self._get_connection_str()

        end_time = time.time() + self._timeout
        last_ex = None
        while end_time > time.time():
            try:
                local_context = cast("XComponentContext", uno.getComponentContext())
                local_factory = local_context.getServiceManager()
                resolver = cast(
                    UnoUrlResolver,
                    local_context.getServiceManager().createInstanceWithContext(
                        "com.sun.star.bridge.UnoUrlResolver", local_context
                    ),
                )
                smgr = resolver.resolve(conn_str)

                props = cast(XPropertySet, smgr.queryInterface(uno.getTypeByName("com.sun.star.beans.XPropertySet")))
                self._ctx = cast(XComponentContext, props.getPropertyValue("DefaultContext"))

                try:
                    bridge_instance = self._get_bridge(local_factory=local_factory, local_ctx=self._ctx)
                except Exception:  # pylint: disable=broad-except
                    bridge_instance = self._get_bridge(local_factory=local_factory, local_ctx=local_context)

                self._bridge_component = cast(
                    XComponent, bridge_instance.queryInterface(uno.getTypeByName("com.sun.star.lang.XComponent"))
                )

                last_ex = None
                break
            except NoConnectException as e:  # pylint: disable=invalid-name
                last_ex = e
                time.sleep(self._conn_try_sleep)

        if last_ex is not None:
            raise last_ex

    def _connect_alternative(self):
        # this method is not currently used.
        # it is an alternative to _connect()
        # This is basically the connect method from Lo.Java
        # it has been tested with local and remote connections and works.
        # Works with Pip and Socket connections.
        conn_str = self._get_connection_identifier()
        # conn_str = "socket,host=localhost,port=2002"

        end_time = time.time() + self._timeout
        last_ex = None
        while end_time > time.time():
            try:
                local_context = cast(XComponentContext, uno.getComponentContext())
                local_factory = local_context.getServiceManager()

                connector = cast(
                    XConnector,
                    local_factory.createInstanceWithContext("com.sun.star.connection.Connector", local_context),
                )

                connection = connector.connect(conn_str)
                # create a bridge to Office via the socket
                bridge_factory = cast(
                    XBridgeFactory,
                    local_factory.createInstanceWithContext("com.sun.star.bridge.BridgeFactory", local_context),
                )

                # create a nameless bridge with no instance provider
                bridge = bridge_factory.createBridge("socketBridgeAD", "urp", connection, None)  # type: ignore
                bridge_component = cast(
                    XBridge, bridge.queryInterface(uno.getTypeByName("com.sun.star.bridge.XBridge"))
                )

                # get the remote service manager
                service_manager = cast(
                    XMultiComponentFactory, bridge_component.getInstance("StarOffice.ServiceManager")
                )

                # retrieve Office's remote component context as a property
                props = cast(
                    XPropertySet, service_manager.queryInterface(uno.getTypeByName("com.sun.star.beans.XPropertySet"))
                )
                default_context = cast(XComponentContext, props.getPropertyValue("DefaultContext"))
                # get the remote interface XComponentContext
                xcc = cast(
                    XComponentContext,
                    default_context.queryInterface(uno.getTypeByName("com.sun.star.uno.XComponentContext")),
                )
                self._ctx = xcc
                self._bridge_component = cast(
                    XComponent, bridge_component.queryInterface(uno.getTypeByName("com.sun.star.lang.XComponent"))
                )

                last_ex = None
                break
            except NoConnectException as e:  # pylint: disable=invalid-name
                last_ex = e
                time.sleep(self._conn_try_sleep)

        if last_ex is not None:
            raise last_ex

    def _popen_from_args(self, args: List[str], shutdown: bool):
        # modified in version 0.12.1
        #  preexec_fn=os.setsid was removed from subprocess.Popen
        # see: https://pastebin.com/tJDwiwvx

        if shutdown:
            if self._platform == SysInfo.PlatformEnum.WINDOWS:
                cmd_str = " ".join(args)
                self._soffice_process_shutdown = subprocess.Popen(cmd_str, shell=True, env=self._environment)
            else:
                self._soffice_process_shutdown = subprocess.Popen(
                    " ".join(args),
                    env=self._environment,
                    # preexec_fn=os.setsid,  # type: ignore
                    shell=True,
                )
        else:
            # start LibreOffice process with python to use with pyuno using subprocess
            # see https://tinyurl.com/y5y66462
            # for unknown reason connection with pipe works fine without shell=True
            # this does not seem to work for socket connections

            # self._soffice_process = subprocess.Popen(
            #     args, env=self._environment, preexec_fn=os.setsid
            # )
            cmd_str = " ".join(args)
            if self._platform == SysInfo.PlatformEnum.WINDOWS:
                self._soffice_process = subprocess.Popen(cmd_str, shell=True, env=self._environment)
            else:
                self._soffice_process = subprocess.Popen(
                    cmd_str,
                    env=self._environment,
                    # preexec_fn=os.setsid,  # type: ignore
                    shell=True,
                )

    def del_working_dir(self):
        """
        Deletes the current working directory of instance.

        This is only applied when caching is used.
        """
        self._cache.del_working_dir()

    def get_soffice_pid(self) -> int | None:
        """
        Gets the pid of soffice

        Returns:
            int: of pid if found; Otherwise, None
        """
        return self._soffice_process.pid if self._soffice_process else None
        # if self._platform == SysInfo.PlatformEnum.WINDOWS:
        #     if self._soffice_process:
        #         return self._soffice_process.pid
        #     return None
        # else:
        #     pid = None
        #     try:
        #         with open(self._pid_file, "r") as f:
        #             pid = f.read()
        #             pid = int(pid)
        #     except Exception:
        #         pid = None
        #     return pid

    def _check_pid(self, pid: int) -> bool:
        """
        Check For the existence of a unix pid.

        Returns:
            bool: True if pid is killed; Otherwise, False
        """
        if pid <= 0:
            return False
        try:
            os.kill(pid, 0)
        except OSError:
            return False
        else:
            return True

    def kill_soffice(self) -> None:
        """
        Attempts to kill instance of soffice created by this instance
        """
        try:
            if self._soffice_process:
                self._soffice_process.kill()
            if self._platform == SysInfo.PlatformEnum.WINDOWS:
                with contextlib.suppress(PermissionError):
                    # this should work without admin privileges.
                    os.system("taskkill /im soffice.bin")
                return

            pid = self.get_soffice_pid()
            if pid is None:
                return None
            # print("pid:", pid)
            if self._check_pid(pid=pid):
                # no SIGLILL on windows.
                os.kill(pid, signal.SIGKILL)  # type: ignore
        except Exception as e:  # pylint: disable=invalid-name
            # print(e)
            raise e

    @property
    def cache(self) -> cache.Cache:
        """
        Gets cache value

        This will always be a Cache instance.
        If no Cache instance is passed into constructor then a default instance is created
        with :py:attr:`Cache.use_cache <.conn.cache.Cache.use_cache>` set to false.
        """
        return self._cache

    @property
    def bridge_component(self) -> XComponent:
        """Gets Bridge Component"""
        return self._bridge_component

    @property
    def soffice(self) -> Path:
        """
        Get Path to LibreOffice soffice. Default is auto discovered.
        """
        return self._connector.soffice

    @property
    def start_office(self) -> bool:
        """Gets if office is to be started. Default is True"""
        return self._connector.start_office

    @property
    def no_restore(self) -> bool:
        """Gets if office is started with norestore Option. Default is True"""
        return self._connector.no_restore

    @property
    def no_first_start_wizard(self) -> bool:
        """Gets if office is started with nofirststartwizard option. Default is True"""
        return self._connector.no_first_start_wizard

    @property
    def no_logo(self) -> bool:
        """Gets if office is started with nologo option. Default is True"""
        return self._connector.no_logo

    @property
    def invisible(self) -> bool:
        """Gets if office is started with invisible option. Default is True"""
        return self._connector.invisible

    @property
    def headless(self) -> bool:
        """Gets/Sets if the connection is made using headless option. Default is False"""
        return self._connector.headless

    @property
    def start_as_service(self) -> bool:
        """
        Gets if office is started as service (StarOffice.Service).
        Default is False
        """
        return self._connector.start_as_service

    @property
    def is_remote(self) -> bool:
        """Gets if connection is connection to remote server. Default is False"""
        return self._connector.remote_connection

    def __del__(self) -> None:
        with contextlib.suppress(Exception):
            self._cache.del_working_dir()


class LoDirectStart(ConnectBase):
    """
    LO Direct Start Connection.

    Used in macros.
    """

    def __eq__(self, other: object) -> bool:
        return isinstance(other, LoDirectStart)

    def connect(self):
        """
        Makes a connection to soffice

        Raises:
            NoConnectException: if unable to obtain a connection to soffice
        """
        self._ctx = uno.getComponentContext()

    def kill_soffice(self) -> None:
        """
        Inherited

        Raises:
            NotImplementedError: Not implement in this class.

        """
        raise NotImplementedError("kill_soffice is not implemented in this child class")

    @property
    def is_remote(self) -> bool:
        """Returns False"""
        return False


class LoPipeStart(LoBridgeCommon):
    """Pipe Start"""

    def __init__(self, connector: connectors.ConnectPipe | None = None, cache_obj: cache.Cache | None = None) -> None:
        if connector is None:
            connector = connectors.ConnectPipe()
        elif not isinstance(connector, connectors.ConnectPipe):
            raise TypeError("connector arg must be ConnectPipe class")

        super().__init__(connector=connector, cache_obj=cache_obj)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, LoPipeStart):
            return False
        local_conn = self._get_connection_str()
        oth_conn = other._get_connection_str()
        return local_conn == oth_conn

    def _get_connection_str(self) -> str:
        return self._connector.get_connection_str()

    def _get_connection_identifier(self) -> str:
        """
        Gets connection identifier

        Such as ``pipe,name="a34rt84y002"``
        """
        return self._connector.get_connection_identifier()

    def connect(self) -> None:
        """
        Connects to office using a pipe

        Raises:
            NoConnectException: If unable to connect
        """
        self._cache.copy_cache_to_profile()
        if self._connector.start_office:
            self._popen()
        try:
            self._connect()
        except NoConnectException as e:  # pylint: disable=invalid-name
            if self._connector.start_office:
                self.kill_soffice()
            raise e
        self._cache.cache_profile()

    def _get_bridge(self, local_factory: XMultiComponentFactory, local_ctx: XComponentContext) -> XBridge:
        connector = cast(
            "XConnector",
            local_factory.createInstanceWithContext("com.sun.star.connection.Connector", local_ctx).queryInterface(
                uno.getTypeByName("com.sun.star.connection.XConnector")
            ),
        )
        bridge_factory = cast(
            "XBridgeFactory",
            local_factory.createInstanceWithContext("com.sun.star.bridge.BridgeFactory", local_ctx).queryInterface(
                uno.getTypeByName("com.sun.star.bridge.XBridgeFactory")
            ),
        )
        conn = connector.connect(f"pipe,name={self.connector.pipe}")
        return bridge_factory.createBridge("PipeBridgeAD", "urp", conn, None)  # type: ignore

    def _popen(self, shutdown=False) -> None:
        # it is important that quotes be placed in the correct place.
        # linux is not fussy on this but in windows it breaks things and you
        # are left wondering what happened.
        # '--accept="socket,host=localhost,port=2002,tcpNoDelay=1;urp;"' THIS WORKS
        # "--accept='socket,host=localhost,port=2002,tcpNoDelay=1;urp;'" THIS FAILS
        # SEE ALSO: https://tinyurl.com/y5y66462
        prefix = "--unaccept=" if shutdown else "--accept="
        soffice_str = str(self._connector.soffice)
        if soffice_str.startswith("flatpak "):
            # special case for flatpak with a space after flatpak
            # not checking directly for flatpak run becuase flatpak --verbose run could be used
            args = [soffice_str]
        else:
            args = [f'"{self._connector.soffice}"']
        self._connector.update_startup_args(args)

        if self._cache.use_cache:
            args.append(f'-env:UserInstallation="{self._cache.user_profile.as_uri()}"')

        args.append(f'{prefix}"pipe,name={self._connector.pipe};urp;"')  # type: ignore

        if self._connector.start_as_service is True:
            args.append("StarOffice.Service")

        self._popen_from_args(args, shutdown)

    @property
    def connector(self) -> connectors.ConnectPipe:
        """Gets the current Connector"""
        return self._connector  # type: ignore

    @property
    def is_remote(self) -> bool:
        """Gets if the connection is remote"""
        return self.connector.remote_connection


class LoSocketStart(LoBridgeCommon):
    """Socket Start"""

    def __init__(
        self, connector: connectors.ConnectSocket | None = None, cache_obj: cache.Cache | None = None
    ) -> None:
        if connector is None:
            connector = connectors.ConnectSocket()
        elif not isinstance(connector, connectors.ConnectSocket):
            raise TypeError("connector arg must be ConnectSocket class")
        super().__init__(connector=connector, cache_obj=cache_obj)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, LoSocketStart):
            return False
        local_conn = self._get_connection_str()
        oth_conn = other._get_connection_str()
        return local_conn == oth_conn

    def _get_connection_str(self) -> str:
        return self._connector.get_connection_str()

    def _get_connection_identifier(self) -> str:
        """
        Gets connection identifier

        Such as ``socket,host=localhost,port=2002``
        """
        return self._connector.get_connection_identifier()

    def connect(self) -> None:
        """
        Connects to office using a pipe

        Raises:
            NoConnectException: If unable to connect
        """
        self._cache.copy_cache_to_profile()
        if self._connector.start_office:
            self._popen()
        try:
            self._connect()
        except NoConnectException as e:  # pylint: disable=invalid-name
            if self._connector.start_office:
                self.kill_soffice()
            raise e
        self._cache.cache_profile()

    def _get_bridge(self, local_factory: XMultiComponentFactory, local_ctx: XComponentContext) -> XBridge:
        connector = cast(
            "XConnector",
            local_factory.createInstanceWithContext("com.sun.star.connection.Connector", local_ctx).queryInterface(
                uno.getTypeByName("com.sun.star.connection.XConnector")
            ),
        )
        bridge_factory = cast(
            "XBridgeFactory",
            local_factory.createInstanceWithContext("com.sun.star.bridge.BridgeFactory", local_ctx).queryInterface(
                uno.getTypeByName("com.sun.star.bridge.XBridgeFactory")
            ),
        )
        conn = connector.connect(f"socket,host={self.connector.host},port={self.connector.port},tcpNoDelay=1")
        return bridge_factory.createBridge("socketBridgeAD", "urp", conn, None)  # type: ignore

    def _popen(self, shutdown=False) -> None:
        # it is important that quotes be placed in the correct place.
        # linux is not fussy on this but in windows it breaks things and you
        # are left wondering what happened.
        # '--accept="socket,host=localhost,port=2002,tcpNoDelay=1;urp;"' THIS WORKS
        # "--accept='socket,host=localhost,port=2002,tcpNoDelay=1;urp;'" THIS FAILS
        # SEE ALSO: https://tinyurl.com/y5y66462
        prefix = "--unaccept=" if shutdown else "--accept="

        soffice_str = str(self._connector.soffice)
        if soffice_str.startswith("flatpak "):
            # special case for flatpak with a space after flatpak
            # not checking directly for flatpak run becuase flatpak --verbose run could be used
            args = [soffice_str]
        else:
            args = [f'"{self._connector.soffice}"']

        self._connector.update_startup_args(args)

        if self._cache.use_cache:
            args.append(f'-env:UserInstallation="{self._cache.user_profile.as_uri()}"')

        args.append(f'{prefix}"socket,host={self._connector.host},port={self._connector.port},tcpNoDelay=1;urp;"')  # type: ignore

        if self._connector.start_as_service is True:
            args.append("StarOffice.Service")

        self._popen_from_args(args, shutdown)

    @property
    def connector(self) -> connectors.ConnectSocket:
        """Gets the current Connector"""
        return self._connector  # type: ignore

    @property
    def is_remote(self) -> bool:
        """Gets if the connection is remote"""
        return self.connector.remote_connection


class LoManager:
    """LO Connection Context Manager"""

    def __init__(
        self,
        connector: connectors.ConnectPipe | connectors.ConnectSocket | None = None,
        cache_obj: cache.Cache | None = None,
    ):
        """
        Context Manager Constructor

        Args:
            connector (connectors.ConnectPipe | connectors.ConnectSocket | None, optional): Connector to connect with. Defaults to ConnectPipe.
            cache_obj (Cache | None, optional): Cache Option. Defaults to None.

        Raises:
            TypeError: If connector is incorrect type.
        """
        if connector is None:
            self._lo = LoPipeStart(cache_obj=cache_obj)
        elif isinstance(connector, connectors.ConnectPipe):
            self._lo = LoPipeStart(connector=connector, cache_obj=cache_obj)
        elif isinstance(connector, connectors.ConnectSocket):
            self._lo = LoSocketStart(connector=connector, cache_obj=cache_obj)
        else:
            raise TypeError("Arg connector is not valid type")

    def __enter__(self) -> LoBridgeCommon:
        self._lo.connect()
        return self._lo

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self._lo.kill_soffice()
