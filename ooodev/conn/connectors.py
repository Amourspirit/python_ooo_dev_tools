# coding: utf-8
from __future__ import annotations
import os
from typing import Dict, Iterable, List, cast
from pathlib import Path
import uuid
from abc import ABC, abstractmethod

from ooodev.utils import paths


class ConnectorBase(ABC):
    @abstractmethod
    def get_connection_str(self) -> str:
        """
        Gets connection string.

        Such as ``uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager``
        """
        ...

    @abstractmethod
    def get_connection_identifier(self) -> str:
        """
        Gets connection identifier

        Such as ``socket,host=localhost,port=2002``
        """
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
        self._env_vars = cast(Dict[str, str], kwargs.get("env_vars", {}))
        self._remote_connection = bool(kwargs.get("remote_connection", False))

        if extended_args := cast(Iterable[str], kwargs.get("extended_args", [])):
            if isinstance(extended_args, str):
                self._extended_args = [extended_args]
            else:
                self._extended_args = list(extended_args)
        else:
            self._extended_args = []

        if soffice := kwargs.get("soffice"):
            # allow empty string or None to be passed
            self._soffice = soffice

    def update_startup_args(self, args: List[str]) -> None:
        # sets does not preserve order
        # preserve order while filtering duplicates
        # Python 3.7 and above guarantee dict order
        args_dict = dict.fromkeys(args)

        if self.no_restore:
            args_dict["--norestore"] = None
        if self.invisible:
            args_dict["--invisible"] = None
        if self.no_restore:
            args_dict["--norestore"] = None
        if self.no_first_start_wizard:
            args_dict["--nofirststartwizard"] = None
        if self.no_logo:
            args_dict["--nologo"] = None
        if self.headless:
            args_dict["--headless"] = None

        if self.extended_args:
            for arg in self.extended_args:
                args_dict[arg] = None

        args.clear()
        # get unique values from dict keys
        args.extend(args_dict.keys())

    @property
    def soffice(self) -> Path:
        """
        Get/Sets Path to LibreOffice soffice. Default is auto discovered.
        """
        try:
            return self._soffice  # type: ignore
        except AttributeError:
            if so := os.environ.get("ODEV_CONN_SOFFICE", None):
                self._soffice = so
            else:
                self._soffice = paths.get_soffice_path()
        return self._soffice  # type: ignore

    @soffice.setter
    def soffice(self, value: Path | str):
        self._soffice = Path(value)

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

    @property
    def env_vars(self) -> Dict[str, str]:
        """Gets/Sets environment variables to be set when starting office"""
        return self._env_vars

    @property
    def extended_args(self) -> List[str]:
        """Extended arguments to be passed to soffice such as ``[--display : 0]``"""
        return self._extended_args

    @property
    def remote_connection(self) -> bool:
        """Specifies if the connection is to a remote server. Default is False"""
        return self._remote_connection

    # endregion startup flags


class ConnectSocket(ConnectorBridgeBase):
    """Connect to LO via socket"""

    def __init__(self, host="localhost", port=2002, **kwargs) -> None:
        """
        Constructor

        Args:
            host (str, optional): Connection host. Defaults to ``localhost``.
            port (int, optional): Connection port. Defaults to ``2002``.

        Keyword Arguments:
            no_restore (bool, optional): Default ``True``
            no_first_start_wizard (bool, optional): Default ``True``
            no_logo (bool, optional): Default ``True``
            invisible (bool, optional): Default ``True``
            headless (bool, optional): Default ``False``
            start_as_service (bool, optional): Default ``False``
            start_office (bool, optional): Default ``True``
            soffice (Path | str, optional): Path to soffice
            env_vars (Dict[str, str], optional): Environment variables to be set when starting office
            extended_args (List[str], optional): Extended arguments to be passed to soffice, such as ``["--display :0"]``.
            remote_connection (bool, optional): Specifies if the connection is to a remote server. Default is False

        Returns:
            None:
        """
        super().__init__(**kwargs)
        self._host = host
        self._port = port

    def get_connection_identifier(self) -> str:
        """
        Gets connection identifier

        Such as ``socket,host=localhost,port=2002``
        """
        return f"socket,host={self.host},port={self.port}"

    def get_connection_str(self) -> str:
        """
        Gets connection string.

        Such as ``uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager``
        """
        identifier = self.get_connection_identifier()
        return f"uno:{identifier};urp;StarOffice.ServiceManager"

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
            pipe (str | None, optional): Name of pipe. Auto generated if None. Defaults to ``None``.

        Keyword Arguments:
            no_restore (bool, optional): Default ``True``
            no_first_start_wizard (bool, optional): Default ``True``
            no_logo (bool, optional): Default ``True``
            invisible (bool, optional): Default ``True``
            headless (bool, optional): Default ``False``
            start_as_service (bool, optional): Default ``False``
            start_office (bool, optional): Default ``True``
            soffice (Path | str, optional): Path to soffice
            env_vars (Dict[str, str], optional): Environment variables to be set when starting office
            extended_args (List[str], optional): Extended arguments to be passed to soffice, such as ``["--display :0"]``.
            remote_connection (bool, optional): Specifies if the connection is to a remote server. Default is False

        Returns:
            None:
        """
        super().__init__(**kwargs)
        self._pipe = uuid.uuid4().hex if pipe is None else pipe

    def get_connection_identifier(self) -> str:
        """
        Gets connection identifier

        Such as ``pipe,name="a34rt84y002"``
        """
        return f"pipe,name={self.pipe}"

    def get_connection_str(self) -> str:
        identifier = self.get_connection_identifier()
        return f"uno:{identifier};urp;StarOffice.ServiceManager"

    @property
    def pipe(self) -> str:
        """
        Gets/Sets pipe used to connect. Default is auto generated hex value
        """
        return self._pipe

    @pipe.setter
    def pipe(self, value: str):
        self._pipe = value
