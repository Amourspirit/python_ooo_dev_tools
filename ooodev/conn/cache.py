# coding: utf-8
from __future__ import annotations
import os
from pathlib import Path
from shutil import copytree
import shutil
import tempfile
from ..utils.type_var import PathOrStr
from ..utils import sys_info


class Cache:
    """Office Profile Cache Manager"""

    def __init__(self, **kwargs) -> None:
        """
        Cache Constructor

        Keyword Args:
            use_cache (bool, optional): Determines if caching is used. Default is True
            cache_path (PathOrStr, optional): Set the path used for caching LO profile data.
                Default is searched for in known locations.
            working_dir (PathOrStr, optional): sets the working dir to use.
                This is the dir that LO profile will be copied to. Defaults to a newly generated temp dir.
        """
        self._use_cache = bool(kwargs.get("use_cache", True))
        self._profile_dir_name = "profile"
        self._profile_cached = False
        self._copy_cache_to_profile_called = False
        cache_path = kwargs.get("cache_path", None)
        if cache_path is not None:
            self.cache_path = cache_path
        working_dir = kwargs.get("working_dir", None)
        if working_dir is not None:
            self.working_dir = working_dir

    def _get_cache_path(self) -> Path | None:
        # see: https://www.howtogeek.com/289587/how-to-find-your-libreoffice-profile-folder-in-windows-macos-and-linux/
        # TODO: the profile paths should be stored in a more global way, at least the 4.
        cache_path = None
        platform = sys_info.SysInfo.get_platform()
        if platform == sys_info.SysInfo.PlatformEnum.LINUX:
            cache_path = Path("~/.config/libreoffice/4").expanduser()
        elif platform == sys_info.SysInfo.PlatformEnum.WINDOWS:
            cache_path = Path(os.getenv("APPDATA"), "LibreOffice", "4")
        elif platform == sys_info.SysInfo.PlatformEnum.MAC:
            cache_path = Path("~/Library/Application Support/LibreOffice/4").expanduser()
        return cache_path

    def cache_profile(self) -> None:
        """
        Copies user profile into cache path.

        Ignored if :py:attr:`~Cache.use_cache` is ``False``
        """
        # copy_cache_to_profile is called before this method
        if self.use_cache is False:
            return
        if self.cache_path is None:
            return
        if self._copy_cache_to_profile_called is False:
            raise Exception("copy_cache_to_profile() must be called before cache_profile()")
        if self._profile_cached is False:
            copytree(self.user_profile, self.cache_path)
        return

    def copy_cache_to_profile(self) -> None:
        """
        Copies cached profile to profile dir set in
        :py:attr:`~.Cache.user_profile`

        Ignored if :py:attr:`~Cache.use_cache` is ``False``
        """
        # this method is called before cache_profile.
        self._copy_cache_to_profile_called = True
        if self.use_cache is False:
            return
        if self.cache_path is None:
            return
        if self.cache_path.exists() and self.cache_path.is_dir():
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
        if self.use_cache:
            if self.working_dir.exists() and self.working_dir.is_dir():
                shutil.rmtree(self.working_dir)

    @property
    def user_profile(self) -> Path:
        """Gets user profile path"""
        try:
            return self._user_profile
        except AttributeError:
            self._user_profile = Path(self.working_dir, self._profile_dir_name)
        return self._user_profile

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
        return self._cache_path

    @cache_path.setter
    def cache_path(self, value: PathOrStr | None):
        self._cache_path = Path(value)

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
