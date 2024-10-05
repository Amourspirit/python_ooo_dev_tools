# coding: utf-8
from __future__ import annotations
import os
from pathlib import Path
from shutil import copytree
import shutil
import tempfile
from typing import Tuple
from ooodev.utils.type_var import PathOrStr
from ooodev.utils import sys_info
from ooodev.cfg import config
from ooodev.io.log.named_logger import NamedLogger


class Cache:
    """Office Profile Cache Manager"""

    def __init__(self, **kwargs) -> None:
        """
        Cache Constructor

        Keyword Args:
            use_cache (bool, optional): Determines if caching is used. Default is True
            cache_path (PathOrStr, optional): Set the path used for caching LO profile data.
                If set to empty string then no profile will be copied and a new profile will be created.
                Default is searched for in known locations.
            working_dir (PathOrStr, optional): sets the working dir to use.
                This is the dir that LO profile will be copied to. Defaults to a newly generated temp dir.
            no_shared_ext (bool, optional): Determines if shared extensions are used.
                If set to True then no shared extensions are disabled for the session.
                Default is False.
        """
        self._log = NamedLogger(self.__class__.__name__)
        self._log.debug("Cache.__init__")
        self._use_cache = bool(kwargs.get("use_cache", True))
        self._no_shared_ext = bool(kwargs.get("no_shared_ext", False))
        self._profile_dir_name = "profile"
        self._no_share_dir_name = "no_share"
        self._profile_cached = False
        cache_path = kwargs.get("cache_path", None)
        if cache_path is not None:
            self.cache_path = cache_path
        working_dir = kwargs.get("working_dir", None)
        if working_dir is not None:
            self.working_dir = working_dir

    def _get_cache_path(self) -> Path | None:
        # this method is only ever called the user does not provide a cache_path
        # see: https://www.howtogeek.com/289587/how-to-find-your-libreoffice-profile-folder-in-windows-macos-and-linux/
        self._log.debug("Cache._get_cache_path()")
        cache_path = None
        platform = sys_info.SysInfo.get_platform()

        def get_path(ver: str) -> Tuple[bool, Path] | Tuple[bool, None]:
            # sourcery skip: merge-nested-ifs
            result = None
            if platform == sys_info.SysInfo.PlatformEnum.LINUX:
                result = Path("~/.config/libreoffice", ver).expanduser()
            elif platform == sys_info.SysInfo.PlatformEnum.WINDOWS:
                result = Path(os.getenv("APPDATA", ""), "LibreOffice", ver)
            elif platform == sys_info.SysInfo.PlatformEnum.MAC:
                result = Path("~/Library/Application Support/LibreOffice/", ver).expanduser()
            if result is not None:
                if result.exists() is False or result.is_dir() is False:
                    result = None
            return (False, None) if result is None else (True, result)

        # lookup profile versions from config
        for s_ver in config.Config().profile_versions:  # type: ignore
            is_valid, cache_path = get_path(s_ver)
            if is_valid:
                break
        self._log.debug(f"Cache._get_cache_path(): {cache_path}")
        return cache_path

    def cache_profile(self) -> None:
        """
        Copies user profile into cache path if it has not already been cached.

        Ignored if :py:attr:`~Cache.use_cache` is ``False``
        """
        # copy_cache_to_profile is called before this method
        if not self.use_cache:
            self._log.debug("Cache.cache_profile(): use_cache is False")
            return
        if not self.cache_path:
            self._log.debug("Cache.cache_profile(): cache_path is None")
            return
        if self._profile_cached is False:
            self._log.debug(
                f"Cache.cache_profile(): copying profile to cache. From: {self.user_profile} To: {self.cache_path}"
            )
            copytree(self.user_profile, self.cache_path)
        return

    def copy_cache_to_profile(self) -> None:
        """
        Copies cached profile to profile dir set in
        :py:attr:`~.Cache.user_profile`

        Ignored if :py:attr:`~Cache.use_cache` is ``False``
        """
        # this method is called before cache_profile.
        if not self.use_cache:
            self._log.debug("Cache.copy_cache_to_profile(): use_cache is False")
            return
        if not self.cache_path:
            self._log.debug("Cache.copy_cache_to_profile(): cache_path is None")
            return
        if self.cache_path.exists() and self.cache_path.is_dir():
            self._log.debug(
                f"Cache.copy_cache_to_profile(): copying cache to profile. From: {self.cache_path} To: {self.user_profile}"
            )
            copytree(self.cache_path, self.user_profile)
            self._profile_cached = True
        else:
            # create the dir.
            # when when cache_profile is called it will cache
            # the profile into this dir.
            os.mkdir(self.user_profile)
            self._profile_cached = False

    def del_working_dir(self):
        """
        Deletes the current working directory of instance.

        Ignored if :py:attr:`~Cache.use_cache` is ``False``
        """
        if self.use_cache and (self.working_dir.exists() and self.working_dir.is_dir()):
            try:
                shutil.rmtree(self.working_dir)
                self._log.debug(f"Cache.del_working_dir(): Deleted working dir: {self.working_dir}")
            except Exception:
                self._log.exception(f"Cache.del_working_dir(): Error deleting working dir: {self.working_dir}.")
        else:
            self._log.debug("Cache.del_working_dir(): working dir does not exist or use_cache is False")

    @property
    def user_profile(self) -> Path:
        """Gets user profile path"""
        try:
            return self._user_profile
        except AttributeError:
            self._user_profile = Path(self.working_dir, self._profile_dir_name)
        return self._user_profile

    @property
    def no_share_path(self) -> Path | None:
        """Gets user No shared extension path"""
        try:
            if self._no_shared_ext is False:
                return None
            return self._no_share_path
        except AttributeError:
            self._no_share_path = Path(self.working_dir, self._no_share_dir_name)
        return self._no_share_path

    @property
    def cache_path(self) -> Path | None:
        """
        Gets/Sets cache path

        The when possible default is obtained by searching in know locations, depending on OS.
        For instance on Linux will search in ``~/.config/libreoffice/4``
        """
        try:
            return self._cache_path
        except AttributeError:
            self._cache_path = self._get_cache_path()
            self._log.debug(f"Cache.cache_path: {self._cache_path}")
        return self._cache_path

    @cache_path.setter
    def cache_path(self, value: PathOrStr | None):
        if not value:
            self._cache_path = None
            return
        self._cache_path = Path(value)  # type: ignore

    @property
    def working_dir(self) -> Path:
        """
        Gets/Sets working dir

        This property determines the dir LO uses to obtain its profile data.
        Default to an auto generated tmp dir.
        """
        try:
            return self._working_dir
        except AttributeError:
            self._working_dir = Path(tempfile.mkdtemp())
            self._log.debug(f"Cache.working_dir: {self._working_dir}")
        return self._working_dir

    @working_dir.setter
    def working_dir(self, value: PathOrStr):
        self._working_dir = Path(value)

    @property
    def use_cache(self) -> bool:
        """Gets/Sets if cache is used. Default is ``True``"""
        return self._use_cache

    @use_cache.setter
    def use_cache(self, value: bool) -> None:
        self._use_cache = value
