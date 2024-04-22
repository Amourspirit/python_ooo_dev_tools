# coding: utf-8
import os
from pathlib import Path
from typing import Union
import uno
from abc import abstractmethod

from ooodev.adapter.util.the_path_settings_comp import ThePathSettingsComp
from ooodev.io.log.named_logger import NamedLogger
from ooodev.meta.constructor_singleton import ConstructorSingleton


class CacheBase(metaclass=ConstructorSingleton):
    """
    Caches files and retrieves cached files.
    Cached file are in a subfolder of system tmp dir.
    """

    def __init__(self, tmp_dir: str = "", lifetime: float = -1) -> None:
        """
        Constructor

        Args:
            tmp_dir (str, optional): Dir name to create in tmp folder. Defaults to 'ooo_uno_tmpl'.
            lifetime (float): Time in seconds that cache is good for.
        """
        uno.fileUrlToSystemPath
        self._ps = ThePathSettingsComp.from_lo()
        t_path = Path(uno.fileUrlToSystemPath(self._ps.temp[0]))
        if tmp_dir:
            self._cache_path = t_path / tmp_dir
            self._cache_path.mkdir(parents=True, exist_ok=True)
        else:
            self._cache_path = t_path

        self._lifetime = lifetime
        self._logger = NamedLogger(self.__class__.__name__)

    @abstractmethod
    def fetch_from_cache(self, filename: Union[str, Path]):
        """
        Fetches file contents from cache if it exist and is not expired

        Args:
            filename (Union[str, Path]): File to retrieve

        Returns:
            Union[str, None]: File contents if retrieved; Otherwise, ``None``
        """

    @abstractmethod
    def save_in_cache(self, filename: Union[str, Path], content):
        """
        Saves file contents into cache

        Args:
            filename (Union[str, Path]): filename to write.
            content (any): Contents to write into file.
        """

    def del_from_cache(self, filename: Union[str, Path]) -> None:
        """
        Deletes a file from cache if it exist

        Args:
            filename (Union[str, Path]): file to delete.
        """
        try:
            f = Path(self.path, filename)
            if os.path.exists(f):
                os.remove(f)
        except Exception as e:
            self.logger.warning(f"Not able to delete file: {filename}, error: {e}")

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
