from __future__ import annotations
from typing import Any
import hashlib
import pickle
from pathlib import Path
from typing import Union
import uno
from ooodev.adapter.util.the_path_settings_comp import ThePathSettingsComp
from ooodev.io.log.named_logger import NamedLogger
from ooodev.meta.constructor_singleton import ConstructorSingleton


class FileChangeAwareCache(metaclass=ConstructorSingleton):
    """
    Singleton Class.
    Caches files and retrieves cached files.
    Cached file are in a subfolder of LibreOffice tmp dir.

    .. versionadded:: 0.52.0
    """

    # region Constructor
    def __init__(self, *, tmp_dir: Union[Path, str] = "", **kwargs: Any) -> None:
        """
        Constructor

        Args:
            tmp_dir (Path, str, optional): Dir name to create in tmp folder. Defaults to ``ooo_uno_tmpl``.
            kwargs (Any): Additional keyword arguments. The arguments are used to create a unique instance of the singleton class.

        Note:
            The cache root temp folder is the LibreOffice temp folder.
        """
        self._tmp_dir = tmp_dir
        ps = ThePathSettingsComp.from_lo()
        t_path = Path(uno.fileUrlToSystemPath(ps.temp[0]))
        if tmp_dir:
            self._cache_path = t_path / tmp_dir
            self._cache_path.mkdir(parents=True, exist_ok=True)
        else:
            self._cache_path = t_path

        self._logger = NamedLogger(self.__class__.__name__)
        self._last_file_path = None
        self._last_key = ""

        self._mod_times = {}

    # endregion Constructor

    # region private methods
    def _get_key(self, file_path: Union[str, Path]) -> str:
        """
        Generates a unique key for the file path

        Args:
            file_path ([str, Path): File path

        Returns:
            str: Unique key
        """
        if not file_path:
            raise ValueError("file_path is required")

        if self._last_file_path == file_path:
            return self._last_key
        self._last_file_path = file_path
        self._last_key = "x" + hashlib.md5(str(file_path).encode("utf-8")).hexdigest()
        return self._last_key

    def _invalidate_if_changed(self, file_path: Union[str, Path], key: str) -> None:
        """
        Invalidates the cache if the file has changed.

        Args:
            file_path (str, Path): File to check for changes
        """

        f = Path(file_path)
        if not f.exists():
            if key in self._mod_times:
                del self._mod_times[key]
            return
        f_stat = f.stat()
        current_mtime = f.stat().st_mtime
        last_mtime = self._mod_times.get(key, None)
        if last_mtime is None:
            # new file
            self._mod_times[key] = current_mtime
            return
        # should not be zero byte file.
        if current_mtime != last_mtime or f_stat.st_size == 0:
            self._mod_times[key] = current_mtime
            self.remove(f)

    # endregion private methods

    # region public methods

    def get(self, file_path: Union[str, Path]) -> Any:
        """
        Fetches file contents from cache if it exist and is not expired

        Args:
            file_path (str, Path): File to retrieve

        Returns:
            Union[object, None]: File contents if retrieved; Otherwise, ``None``
        """

        key = self._get_key(file_path)
        cache_file_name = key + ".pkl"

        self._invalidate_if_changed(file_path, key)

        f = Path(self.path, cache_file_name)
        if not f.exists():
            return None

        try:
            # Open the file in binary mode
            with open(f, "rb") as file:
                # Call load method to deserialize
                content = pickle.load(file)
            return content
        # except IOError:
        #     return None
        except Exception:
            self.logger.exception("Error reading file: %s", f)
            return None

    def put(self, file_path: Union[str, Path], content: Any):
        """
        Saves file contents into cache

        Args:
            file_path (str, Path): filename to write.
            content (Any): Contents to write into file.
        """
        key = self._get_key(file_path)

        if isinstance(file_path, str):
            file_path = Path(file_path)

        cache_file_name = key + ".pkl"
        f = Path(self.path, cache_file_name)
        if f.exists():
            f.unlink()
        with open(f, "wb") as file:
            pickle.dump(content, file)

        current_mtime = file_path.stat().st_mtime
        self._mod_times[key] = current_mtime

    def remove(self, file_path: Union[str, Path]) -> None:
        """
        Deletes a file from cache if it exist

        Args:
            file_path (str, Path): file to delete.
        """
        key = self._get_key(file_path)
        try:
            cache_file_name = key + ".pkl"
            f = Path(self.path, cache_file_name)
            if f.exists():
                f.unlink()
                self._logger.debug("Deleted file: %s", f)
        except Exception as e:
            self.logger.warning("Not able to delete file: %s, error: %s", file_path, e)

    # endregion public methods

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
        return f"<{self.__class__.__name__}(tmp_dir={self._tmp_dir})>"

    # endregion Dunder Methods

    # region Properties
    @property
    def path(self) -> Path:
        """Gets cache path"""
        return self._cache_path

    @property
    def logger(self) -> NamedLogger:
        """Gets logger"""
        return self._logger

    # endregion Properties
