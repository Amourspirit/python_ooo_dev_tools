from __future__ import annotations
from typing import Any
import pickle
import time
from pathlib import Path
from typing import Union
from ooodev.utils.cache.file_cache.cache_base import CacheBase


class PickleCache(CacheBase):
    """
    Caches files and retrieves cached files.
    Cached file are in a subfolder of system tmp dir.
    """

    def get(self, filename: Union[str, Path]) -> Any:
        """
        Fetches file contents from cache if it exist and is not expired

        Args:
            filename (Union[str, Path]): File to retrieve

        Returns:
            Union[object, None]: File contents if retrieved; Otherwise, ``None``
        """
        f = Path(self.path, filename)
        if not f.exists():
            return None
        if self.can_expire:
            f_stat = f.stat()
            if f_stat.st_size == 0:
                # should not be zero byte file.
                try:
                    self.remove(f)
                except Exception as e:
                    self.logger.warning(f"Not able to delete 0 byte file: {filename} error: {e}")
                return None
            ti_m = f_stat.st_mtime
            age = time.time() - ti_m
            if age >= self.seconds:
                return None

        try:
            # Open the file in binary mode
            with open(f, "rb") as file:
                # Call load method to deserialize
                content = pickle.load(file)
            return content
        except IOError:
            return None
        except Exception as e:
            self.logger.error(e, exc_info=True)
            raise e

    def put(self, filename: Union[str, Path], content: Any):
        """
        Saves file contents into cache

        Args:
            filename (Union[str, Path]): filename to write.
            content (Any): Contents to write into file.
        """
        f = Path(self.path, filename)
        with open(f, "wb") as file:
            pickle.dump(content, file)
