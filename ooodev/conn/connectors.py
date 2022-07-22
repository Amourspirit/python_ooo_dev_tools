# coding: utf-8
from __future__ import annotations
from typing import List
from pathlib import Path
import uuid
from abc import ABC, abstractmethod

from ..utils import paths


class ConnectorBase(ABC):
    @abstractmethod
    def get_connnection_str(self) -> str:
        ...


class ConnectorBridgeBase(ConnectorBase):
    def __init__(self, **kwargs) -> None:
        self._no_restore = bool(kwargs.get("no_restore", True))
        self._no_first_start_wizard = bool(kwargs.get("no_first_start_wizard", True))
        self._no_logo = bool(kwargs.get("no_logo", True))
        self._invisible = bool(kwargs.get("invisible", True))
        self._headless = bool(kwargs.get("headless", False))
        self._start_as_service = bool(kwargs.get("start_as_service", False))
        self._start_office = bool(kwargs.get("start_office", True))

        if "soffice" in kwargs:
            self.soffice = kwargs["soffice"]

    def update_startup_args(self, args: List[str]) -> None:
        if self.no_restore:
            args.append("--norestore")
        if self.invisible:
            args.append("--invisible")
        if self.no_restore:
            args.append("--norestore")
        if self.no_first_start_wizard:
            args.append("--nofirststartwizard")
        if self.no_logo:
            args.append("--nologo")
        if self.headless:
            args.append("--headless")

    @property
    def soffice(self) -> Path:
        """
        Get/Sets Path to LibreOffice soffice. Default is auto discovered.
        """
        try:
            return self._soffice
        except AttributeError:
            self._soffice = paths.get_soffice_path()
        return self._soffice

    # region startup flags
    @property
    def start_office(self) -> bool:
        """Gets/Sets if office is to be started. Default is True"""
        return self._start_office

    @start_office.setter
    def start_office(self, value: bool):
        self._start_office = value

    @property
    def no_restore(self) -> bool:
        """Gets/Sets if office is started with norestore Option. Default is True"""
        return self._no_restore

    @no_restore.setter
    def no_restore(self, value: bool):
        self._no_restore = value

    @property
    def no_first_start_wizard(self) -> bool:
        """Gets/Sets if office is started with nofirststartwizard option. Default is True"""
        return self._no_first_start_wizard

    @no_first_start_wizard.setter
    def no_first_start_wizard(self, value: bool):
        self._no_first_start_wizard = value

    @property
    def no_logo(self) -> bool:
        """Gets/Sets if office is started with nologo option. Default is True"""
        return self._no_logo

    @no_logo.setter
    def no_logo(self, value: bool):
        self._no_logo = value

    @property
    def invisible(self) -> bool:
        """Gets/Sets if office is started with invisible option. Default is True"""
        return self._invisible

    @invisible.setter
    def invisible(self, value: bool):
        self._invisible = value

    @property
    def headless(self) -> bool:
        """Gets/Sets if the connection is made using headless option. Default is False"""
        return self._headless

    @headless.setter
    def headless(self, value: bool):
        self._headless = value

    @property
    def start_as_service(self) -> bool:
        """
        Gets/Sets if office is started as service (StarOffice.Service).
        Default is False
        """
        return self._start_as_service

    @start_as_service.setter
    def start_as_service(self, value: bool):
        self._start_as_service = value

    # endregion startup flags


class ConnectSocket(ConnectorBridgeBase):
    """Connect to LO via socket"""
    def __init__(self, host="localhost", port=2002, **kwargs) -> None:
        """
        Constructor

        Args:
            host (str, optional): Connection host. Defaults to "localhost".
            port (int, optional): Connection port. Defaults to 2002.
        
        Any property value can be set by passing it into constructor.
        e.g. ``ConnectSocket(no_logo=True)``
        """
        super().__init__(**kwargs)
        self._host = host
        self._port = port

    def get_connnection_str(self) -> str:
        identifier = f"socket,host={self.host},port={self.port}"
        conn_str = f"uno:{identifier};urp;StarOffice.ServiceManager"
        return conn_str

    @property
    def host(self) -> str:
        """
        Gets/Sets host. Default ``localhost``
        """
        return self._host

    @host.setter
    def host(self, value: str):
        self._host = value

    @property
    def port(self) -> int:
        """
        Gets/Sets port. Default ``2002``
        """
        return self._port

    @port.setter
    def port(self, value: int):
        self._port = value


class ConnectPipe(ConnectorBridgeBase):
    """Connect to LO via pipe"""
    def __init__(self, pipe: str | None = None, **kwargs) -> None:
        """
        Constructor

        Args:
            pipe (str | None, optional): name of pipe. Auto generated if None. Defaults to None.
        
        Any property value can be set by passing it into constructor.
        e.g. ``ConnectPipe(no_logo=True)``
        """
        super().__init__(**kwargs)
        if pipe is None:
            self._pipe = uuid.uuid4().hex
        else:
            self._pipe = pipe

    def get_connnection_str(self) -> str:
        identifier = f"pipe,name={self.pipe}"
        conn_str = f"uno:{identifier};urp;StarOffice.ServiceManager"
        return conn_str

    @property
    def pipe(self) -> str:
        """
        Gets/Sets pipe used to connect. Default is auto generated hex value
        """
        return self._pipe

    @pipe.setter
    def pipe(self, value: str):
        self._pipe = value
