# coding: utf-8
from __future__ import annotations
import os
import time
from abc import ABC, abstractmethod
import subprocess
import signal
from typing import List, TYPE_CHECKING, cast
import time
import shutil
import uno
from com.sun.star.connection import NoConnectException
from . import connectors
from . import cache
from ..utils.sys_info import SysInfo


if TYPE_CHECKING:
    from com.sun.star.beans import XPropertySet
    from com.sun.star.bridge import XBridgeFactory
    from com.sun.star.bridge import XBridge
    from com.sun.star.connection import XConnector
    from com.sun.star.lang import XComponent
    from com.sun.star.lang import XMultiComponentFactory
    from com.sun.star.uno import XComponentContext


class ConnectBase(ABC):
    """Base Abstract Class for all connections to LO"""

    def __init__(self):
        # https://tinyurl.com/yb897bxw
        # https://tinyurl.com/ybk7zqcg

        # start openoffice process with python to use with pyuno using subprocess
        # see https://tinyurl.com/y5y66462
        self._ctx: XComponentContext = None

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
        return self._ctx


class LoBridgeCommon(ConnectBase):
    """Base Abstract Class for LoSocketStart and LoPipeStart"""

    def __init__(self, cache_obj: cache.Cache | None):
        super().__init__()
        self._soffice_process = None
        self._bridge_component = None
        self._platform = SysInfo.get_platform()
        self._environment = os.environ
        self._timeout = 30.0
        self._conn_try_sleep = 0.2
        if cache_obj is None:
            self._cache = cache.Cache(use_cache=False)
        else:
            self._cache = cache_obj
        if self._cache.use_cache:
            self._environment["TMPDIR"] = str(self._cache.working_dir)

    @abstractmethod
    def _get_connection_str(self) -> str:
        ...

    @abstractmethod
    def _get_bridge(self, local_factory: XMultiComponentFactory, local_ctx: XComponentContext) -> XBridge:
        ...

    def _connect(self):
        conn_str = self._get_connection_str()

        end_time = time.time() + self._timeout
        last_ex = None
        while end_time > time.time():
            try:
                localContext = cast("XComponentContext", uno.getComponentContext())
                # resolver = cast("XMultiComponentFactory",localContext.ServiceManager.createInstanceWithContext(
                #     "com.sun.star.bridge.UnoUrlResolver", localContext
                # ))

                localFactory = localContext.getServiceManager()

                bridge = self._get_bridge(local_factory=localFactory, local_ctx=localContext)

                self._bridge_component = bridge.queryInterface(uno.getTypeByName("com.sun.star.lang.XComponent"))

                # smgr = resolver.resolve(conn_str)
                smgr = cast(
                    "XMultiComponentFactory",
                    bridge.getInstance("StarOffice.ServiceManager").queryInterface(
                        uno.getTypeByName("com.sun.star.lang.XMultiComponentFactory")
                    ),
                )
                props = cast("XPropertySet", smgr.queryInterface(uno.getTypeByName("com.sun.star.beans.XPropertySet")))

                self._ctx = props.getPropertyValue("DefaultContext")
                last_ex = None
                break
            except NoConnectException as e:
                last_ex = e
                time.sleep(self._conn_try_sleep)

        if last_ex is not None:
            raise last_ex

    def _popen_from_args(self, args: List[str], shutdown: bool):
        if shutdown == True:
            if self._platform == SysInfo.PlatformEnum.WINDOWS:
                cmd_str = " ".join(args)
                self._soffice_process_shutdown = subprocess.Popen(cmd_str, shell=True, env=self._envirnment)
            else:
                self._soffice_process_shutdown = subprocess.Popen(
                    " ".join(args),
                    env=self._envirnment,
                    preexec_fn=os.setsid,
                    shell=True,
                )
        else:
            # start LibreOffice process with python to use with pyuno using subprocess
            # see https://tinyurl.com/y5y66462
            # for unknown reason connection with pipe works fine without shell=True
            # this does not seem to work for socket connections

            # self._soffice_process = subprocess.Popen(
            #     args, env=self._envirnment, preexec_fn=os.setsid
            # )
            cmd_str = " ".join(args)
            if self._platform == SysInfo.PlatformEnum.WINDOWS:
                self._soffice_process = subprocess.Popen(cmd_str, shell=True, env=self._environment)
            else:
                self._soffice_process = subprocess.Popen(
                    cmd_str,
                    env=self._environment,
                    preexec_fn=os.setsid,
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
        if self._platform == SysInfo.PlatformEnum.WINDOWS:
            if self._soffice_process:
                return self._soffice_process.pid
            return None
        else:
            pid = None
            try:
                with open(self._pid_file, "r") as f:
                    pid = f.read()
                    pid = int(pid)
            except:
                pid = None
            return pid

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
                try:
                    # this should work without admin privileges.
                    os.system("taskkill /im soffice.bin")
                except PermissionError:
                    # Not able to terminate.
                    # Windows issue, Needs to be run a admin.
                    pass
                return

            pid = self.get_soffice_pid()
            if pid is None:
                return None
            # print("pid:", pid)
            if self._check_pid(pid=pid):
                # no SIGLILL on windows.
                os.kill(pid, signal.SIGKILL)
        except Exception as e:
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
        return self._bridge_component

    def __del__(self) -> None:
        try:
            self._cache.del_working_dir()
        except Exception:
            pass


class LoDirectStart(ConnectBase):
    def __init__(self):
        """
        Constructor
        """
        super().__init__()

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


class LoPipeStart(LoBridgeCommon):
    def __init__(self, connector: connectors.ConnectPipe | None = None, cache_obj: cache.Cache | None = None) -> None:
        super().__init__(cache_obj=cache_obj)
        if connector is None:
            self._connector = connectors.ConnectPipe()
        else:
            if not isinstance(connector, connectors.ConnectPipe):
                raise TypeError("connector arg must be ConnectPipe class")
            self._connector = connector

    def _get_connection_str(self) -> str:
        return self._connector.get_connnection_str()

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
        except NoConnectException as e:
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
        bridgeFactory = cast(
            "XBridgeFactory",
            local_factory.createInstanceWithContext("com.sun.star.bridge.BridgeFactory", local_ctx).queryInterface(
                uno.getTypeByName("com.sun.star.bridge.XBridgeFactory")
            ),
        )
        xconn = connector.connect(f"pipe,name={self.connector.pipe}")
        bridge = bridgeFactory.createBridge("PipeBridgeAD", "urp", xconn, None)
        return bridge

    def _popen(self, shutdown=False) -> None:
        # it is important that quotes be placed in the correct place.
        # linux is not fussy on this but in windows it breaks things and you
        # are left wondering what happened.
        # '--accept="socket,host=localhost,port=2002,tcpNoDelay=1;urp;"' THIS WORKS
        # "--accept='socket,host=localhost,port=2002,tcpNoDelay=1;urp;'" THIS FAILS
        # SEE ALSO: https://tinyurl.com/y5y66462
        if shutdown == True:
            prefix = "--unaccept="
        else:
            prefix = "--accept="

        args = [f'"{self._connector.soffice}"']
        self._connector.update_startup_args(args)

        if self._cache.use_cache:
            args.append(f'-env:UserInstallation="{self._cache.user_profile.as_uri()}"')

        args.append(f'{prefix}"pipe,name={self._connector.pipe};urp;"')

        if self._connector.start_as_service is True:
            args.append("StarOffice.Service")

        self._popen_from_args(args, shutdown)

    @property
    def connector(self) -> connectors.ConnectPipe:
        """Gets the current Connector"""
        return self._connector


class LoSocketStart(LoBridgeCommon):
    def __init__(
        self, connector: connectors.ConnectSocket | None = None, cache_obj: cache.Cache | None = None
    ) -> None:
        super().__init__(cache_obj=cache_obj)
        if connector is None:
            self._connector = connectors.ConnectSocket()
        else:
            if not isinstance(connector, connectors.ConnectSocket):
                raise TypeError("connector arg must be ConnectSocket class")
            self._connector = connector

    def _get_connection_str(self) -> str:
        return self._connector.get_connnection_str()

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
        except NoConnectException as e:
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
        bridgeFactory = cast(
            "XBridgeFactory",
            local_factory.createInstanceWithContext("com.sun.star.bridge.BridgeFactory", local_ctx).queryInterface(
                uno.getTypeByName("com.sun.star.bridge.XBridgeFactory")
            ),
        )
        xconn = connector.connect(f"socket,host={self.connector.host},port={self.connector.port},tcpNoDelay=1")
        bridge = bridgeFactory.createBridge("socketBridgeAD", "urp", xconn, None)
        return bridge

    def _popen(self, shutdown=False) -> None:
        # it is important that quotes be placed in the correct place.
        # linux is not fussy on this but in windows it breaks things and you
        # are left wondering what happened.
        # '--accept="socket,host=localhost,port=2002,tcpNoDelay=1;urp;"' THIS WORKS
        # "--accept='socket,host=localhost,port=2002,tcpNoDelay=1;urp;'" THIS FAILS
        # SEE ALSO: https://tinyurl.com/y5y66462
        if shutdown == True:
            prefix = "--unaccept="
        else:
            prefix = "--accept="

        args = [f'"{self._connector.soffice}"']
        self._connector.update_startup_args(args)

        if self._cache.use_cache:
            args.append(f'-env:UserInstallation="{self._cache.user_profile.as_uri()}"')

        args.append(f'{prefix}"socket,host={self._connector.host},port={self._connector.port},tcpNoDelay=1;urp;"')

        if self._connector.start_as_service is True:
            args.append("StarOffice.Service")

        self._popen_from_args(args, shutdown)

    @property
    def connector(self) -> connectors.ConnectSocket:
        """Gets the current Connector"""
        return self._connector


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
