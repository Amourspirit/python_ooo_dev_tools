# coding: utf-8
import os
from pathlib import Path
from typing import Any, Union
import uno
from abc import abstractmethod

from ooodev.adapter.util.the_path_settings_comp import ThePathSettingsComp
from ooodev.io.log.named_logger import NamedLogger
from ooodev.meta.constructor_singleton import ConstructorSingleton


class CacheBase(metaclass=ConstructorSingleton):
    """
    Caches files and retrieves cached files.
    Cached file are in ``ooo_uno_tmpl` subdirectory of LibreOffice tmp dir.
    """

    def __init__(self, tmp_dir: Union[Path, str] = "", lifetime: float = -1, **kwargs: Any) -> None:
        """
        Constructor

        Args:
            tmp_dir (Path, str, optional): Dir name to create in tmp folder.
            lifetime (float): Time in seconds that cache is good for.
            kwargs (Any): Additional keyword arguments. The arguments are used to create a unique instance of the singleton class.

        Note:
            The cache root temp folder is the LibreOffice temp folder.
        """
        self._ps = ThePathSettingsComp.from_lo()
        t_path = Path(uno.fileUrlToSystemPath(self._ps.temp[0]))
        self._tmp_dir = tmp_dir
        if tmp_dir:
            self._cache_path = t_path / tmp_dir
            self._cache_path.mkdir(parents=True, exist_ok=True)
        else:
            self._cache_path = t_path

        self._lifetime = lifetime
        self._logger = NamedLogger(self.__class__.__name__)

    @abstractmethod
    def get(self, filename: str) -> Any:
        """
        Fetches file contents from cache if it exist and is not expired

        Args:
            filename (str): File to retrieve

        Returns:
            Any: File contents if retrieved; Otherwise, ``None``
        """
        ...

    @abstractmethod
    def put(self, filename: str, content: Any) -> None:
        """
        Saves file contents into cache

        Args:
            filename (str): filename to write.
            content (Any): Contents to write into file.
        """
        ...

    def remove(self, filename: Union[str, Path]) -> None:
        """
        Deletes a file from cache if it exist

        Args:
            filename (Union[str, Path]): file to delete.
        """
        if not filename:
            raise ValueError("filename is required")
        try:
            f = Path(self.path, filename)
            if os.path.exists(f):
                os.remove(f)
        except Exception as e:
            self.logger.warning("Not able to delete file: %s, error: %s", filename, e)

    # region Dunder Methods
    def __contains__(self, key: Any) -> bool:
        return self.get(key) is not None

    def __getitem__(self, key: Any) -> Any:
        return self.get(key)

    def __setitem__(self, key: Any, value: Any) -> None:
        self.put(key, value)

    def __delitem__(self, key: Any) -> None:
        self.remove(key)

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}(tmp_dir={self._tmp_dir},  lifetime={self._lifetime})>"

    # endregion Dunder Methods

    # region Properties
    @property
    def seconds(self) -> float:
        """Gets/Sets cache time in seconds"""
        return self._lifetime

    @seconds.setter
    def seconds(self, value: float) -> None:
        self._lifetime = float(value)

    @property
    def can_expire(self) -> bool:
        """Gets/Sets cache expiration"""
        return self._lifetime > 0

    @property
    def path(self) -> Path:
        """Gets cache path"""
        return self._cache_path

    @property
    def path_settings(self) -> ThePathSettingsComp:
        """Gets path settings"""
        return self._ps

    @property
    def logger(self) -> NamedLogger:
        """Gets logger"""
        return self._logger

    # endregion Properties
