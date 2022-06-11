# coding: utf-8
from __future__ import annotations
import os
import sys
from abc import ABC, abstractmethod, abstractstaticmethod
from pathlib import Path
import subprocess
import tempfile
import signal
from typing import Union, TYPE_CHECKING
import uuid
import time
import shutil
from distutils.dir_util import copy_tree
import uno
from . import paths
from com.sun.star.connection import NoConnectException

if TYPE_CHECKING:
    from com.sun.star.script.provider import XScriptContext


class ConnectBase(ABC):
    def __init__(self, **kwargs):
        # https://tinyurl.com/yb897bxw
        # https://tinyurl.com/ybk7zqcg

        # start openoffice process with python to use with pyuno using subprocess
        # see https://tinyurl.com/y5y66462
        self._is_nt = sys.platform == "win32"
        self._script_content: XScriptContext = None

    @abstractmethod
    def connect(self):
        """
        Makes a connection to soffice

        Raises:
            NoConnectException: if unable to obtain a connection to soffice
        """
        ...
   
    @property
    def desktop(self):
        try:
            return self._desktop
        except AttributeError:
            self._desktop = self.service_manager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx
            )
        return self._desktop

    @property
    def service_manager(self):
        try:
            return self._service_manager
        except AttributeError:
            self._service_manager = self.ctx.getServiceManager()
        return self._service_manager

    @property
    def profile_path_user(self) -> Path:
        """
        Gets the profile user path for the current instance. Usually working_dir + '/profile/user'
        """
        try:
            return self._profile_path_user
        except AttributeError:
            self._profile_path_user = Path(self.profile_path, self._profile_user_name)
        return self._profile_path_user

    @property
    def ctx(self) -> XScriptContext:
        return self._script_content


class LoManager:
    def __init__(self, use_pipe: bool = True, **kwargs):
        """
        Constructor

        Keyword Arguments:
            soffice_path (str, optional): the path to soffice: Default ``/usr/bin/soffice``
            working_dir (str, optional): the working directory of LO. This is the path that profile will be copied into. Defaults to a tmp directory
            cache_path (str, optional): the path where the current profile resides to copy into the LO
                workding dir for current instance of LO
        """
        if use_pipe:
            self._lo = LoPipeStart(**kwargs)
        else:
            self._lo = LoSocketStart(**kwargs)

    def __enter__(self) -> ConnectBase:
        self._lo.connect()
        return self._lo

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self._lo.kill_soffice()


class LoBridgeCommon(ConnectBase):
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self._soffice_install_path = None
        self.profile_cached = False
        self._soffice_process_shutdown = None
        self._profile_dir_name = "profile"
        self._profile_user_name = "user"       
        self._invisible = True
        self._headless = False
        self._start_as_service = False
        self._start_soffice = bool(kwargs.get("start_soffice", True))
        self._soffice_path = self._get_soffice_exe()
        self._working_dir: Path = None
        self._cache_path: Path = None
        self._use_caching = bool(kwargs.get('use_caching', False))

        if self._is_nt:
            # os.getenv('APPDATA')
            default_profile_path = Path(
                os.getenv("APPDATA"), "LibreOffice", "4"
            ).as_posix()
        else:
            default_profile_path = Path(
                os.environ["HOME"], ".config", "libreoffice", "4"
            ).as_posix()



        self._soffice_install_path = str(kwargs.get("soffice_path", self._soffice_install_path))
        working_dir = kwargs.get("working_dir")
        if working_dir is None:
            working_dir = Path(tempfile.mkdtemp())
        else:
            working_dir = Path(working_dir)
        self._working_dir = working_dir
        cache_path = kwargs.get("cache_path", None)
        if cache_path is None:
            cache_path = Path(os.getenv("XDG_CACHE_DIR", default_profile_path))
        else:
            cache_path = Path(cache_path)
        self._cache_path = cache_path
        
        self._user_profie = Path(self._working_dir, self._profile_dir_name)

        self._envirnment = os.environ
        self._envirnment["TMPDIR"] = str(self._working_dir)
        self._conn_try_count = 150
        self._conn_try_time_sleep = 0.2
        if self._is_nt:
            self._pid_file = None
        else:
            self._pid_file = os.path.join(self._working_dir, "soffice.pid")

    @abstractstaticmethod
    def _popen(self, **kwargs) -> None:
        ...

    @abstractstaticmethod
    def _connect(self) -> None:
        ...

    def connect(self):
        """
        Makes a connection to soffice

        Raises:
            NoConnectException: if unable to obtain a connection to soffice
        """
        self._copy_cache_to_profile()
        if self._start_soffice:
            self._popen()
        try:
            self._connect()
        except NoConnectException as e:
            if self._start_soffice:
                self.kill_soffice()
            raise e
        self._cache_current_profile()

    def get_soffice_pid(self) -> Union[int, None]:
        """
        Gets the pid of soffice

        Returns:
            int: of pid if found; Otherwise, `None`
        """
        if self._is_nt:
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
    def del_working_dir(self):
        """
        Deletes the current working directory of instance
        """
        if os.path.exists(self.working_dir):
            shutil.rmtree(self.working_dir)

    def kill_soffice(self) -> None:
        """
        Attempts to kill instance of soffice created by this instance
        """
        try:
            # self._popen(shutdown=True)
            if self._soffice_process_shutdown:
                self._soffice_process_shutdown.kill()
            if self._soffice_process:
                self._soffice_process.kill()
            if self._is_nt:
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
            if self.check_pid(pid=pid):
                # not SIGLILL on windows.
                os.kill(pid, signal.SIGKILL)
            # cls.lo_start.soffice_process.kill()
        except Exception as e:
            # print(e)
            raise e

    def _popen(self, **kwargs):
        # it is important that quotes be placed in the correct place.
        # linux is not fussy on this but in windows it breaks things and you
        # are left wondering what happened.
        # '--accept="socket,host=localhost,port=2002,tcpNoDelay=1;urp;"' THIS WORKS
        # "--accept='socket,host=localhost,port=2002,tcpNoDelay=1;urp;'" THIS FAILS
        # SEE ALSO: https://tinyurl.com/y5y66462
        args = [
            self._soffice_path,
            "--pidfile=" + self._pid_file,
            f'--accept="pipe,name={self._pipe_name};urp;"',
            "--norestore",
            "--invisible",
        ]
        if self._use_caching:
            args.append(f'-env:UserInstallation="file:///{self._user_profie}"')
        str_args = " ".join(args)
        if self._is_nt:
            self._soffice_process = subprocess.Popen(
                    str_args, shell=True, env=self._envirnment
                )
        else:
            self._soffice_process = subprocess.Popen(
                str_args, env=self._envirnment, preexec_fn=os.setsid, shell=True
            )

    def check_pid(self, pid: int) -> bool:
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

    def _get_soffice_exe(self) -> str:
        if self._is_nt:
            soffice = "soffice.exe"
            p_sf = Path(self.soffice_install_path, "program", soffice)
            if not p_sf.exists():
                raise FileNotFoundError(f"LibreOffice '{p_sf}' not found.")
            if not p_sf.is_file():
                raise IsADirectoryError(f"LibreOffice '{p_sf}' is not a file.")
            return str(p_sf)
        else:
            soffice = "soffice"
            # search system path
            s = shutil.which(soffice)
            p_sf = None
            if s is not None:
                # expect '/usr/bin/soffice'
                if os.path.islink(s):
                    p_sf = Path(os.path.realpath(s))
                else:
                    p_sf = Path(s)
            if p_sf is None:
                p_sf = Path("/usr/bin/soffice")
            if not p_sf.exists():
                raise FileNotFoundError(f"LibreOffice '{p_sf}' not found.")
            if not p_sf.is_file():
                raise IsADirectoryError(f"LibreOffice '{p_sf}' is not a file.")
            return str(p_sf)

    def _cache_current_profile(self):
        if self._use_caching is False:
            return
        if self.profile_cached == False:
            copy_tree(self._user_profie, self._cache_path)

    def _copy_cache_to_profile(self):
        if self._use_caching is False:
            return
        if os.path.isdir(self._cache_path):
            # shutil.copytree(cacheDir, userProfile)
            copy_tree(str(self._cache_path), str(self._user_profie))
            self.profile_cached = True
        else:
            os.mkdir(self._user_profie)
            self.profile_cached = False

    @property
    def soffice_install_path(self) -> Path:
        if self._soffice_install_path is None:
            self._soffice_install_path = paths.get_soffice_install_path()
        return self._soffice_install_path
    
    @property
    def soffice_process(self):
        return self._soffice_process

    @property
    def working_dir(self) -> Path:
        return self._working_dir

    @property
    def port(self) -> int:
        return self._port

    @property
    def host(self) -> str:
        return self._host

    @property
    def profile_path(self) -> Path:
        """
        Gets the profile path for the current instance. Usually working_dir + '/profile'
        """
        return self._user_profie


class LoPipeStart(LoBridgeCommon):
    def __init__(self, **kwargs):
        """
        Constructor

        Keyword Arguments:
            soffice_path (str, optional): the path to soffice: Default ``/usr/bin/soffice``
            working_dir (str, optional): the working directory of LO. This is the path that profile will be copied into. Defaults to a tmp directory
            cache_path (str, optional): the path where the current profile resides to copy into the LO
                workding dir for current instance of LO
            start_soffice (bool, optional): If True soffice will be started a server that can be connected to. Default True
        """
        super().__init__(**kwargs)
        self._pipe_name = uuid.uuid4().hex

    

    def _connect(self):

        identifier = f"pipe,name={self._pipe_name}"
        conn_str = f"uno:{identifier};urp;StarOffice.ServiceManager"
        for i in range(self._conn_try_count):
            try:
                localContext = uno.getComponentContext()
                resolver = localContext.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", localContext
                )
                smgr = resolver.resolve(conn_str)
                self._script_content = smgr.getPropertyValue("DefaultContext")
                # self._script_content = resolver.resolve(
                #     "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % self._pipe_name)
                break
            except NoConnectException as e:
                time.sleep(self._conn_try_time_sleep)
                if i == self._conn_try_count - 1:
                    raise e


class LoSocketStart(LoBridgeCommon):
    def __init__(self, **kwargs):
        """
        Constructor

        Keyword Arguments:
            soffice_path (str, optional): the path to soffice: Default System dependent.
            working_dir (str, optional): the working directory of LO. This is the path that profile will be copied into. Defaults to a tmp directory
            cache_path (str, optional): the path where the current profile resides to copy into the LO
                workding dir for current instance of LO
            host (str, optional): host to connect. Default localhost
            port (int, optional): connection port. Default 2002
            start_soffice (bool, optional): If True soffice will be started a server that can be connected to. Default True
        """
        super().__init__(**kwargs)
        self._host = str(kwargs.get("host", "localhost"))
        self._port = int(kwargs.get("port", 2002))

    def _popen(self, shutdown=False):
        # it is important that quotes be placed in the correct place.
        # linux is not fussy on this but in windows it breaks things and you
        # are left wondering what happened.
        # '--accept="socket,host=localhost,port=2002,tcpNoDelay=1;urp;"' THIS WORKS
        # "--accept='socket,host=localhost,port=2002,tcpNoDelay=1;urp;'" THIS FAILS
        # SEE ALSO: https://tinyurl.com/y5y66462
        args = [
            f'"{self._soffice_path}"',
        ]
        if self._use_caching:
            args.append(f'-env:UserInstallation="file:///{self._user_profie.as_posix()}"')
        if self._is_nt is False:
            args.append(f"--pidfile={self._pid_file}")
        args.append("--norestore")
        args.append("--nofirststartwizard")
        args.append("--nologo")
        if self._invisible:
            args.append("--invisible")
        if shutdown == True:
            prefix = "--unaccept="
        else:
            prefix = "--accept="
        conn_suf = f"socket,host={self._host},port={self._port},tcpNoDelay=1;urp;"
        if self._start_as_service == True:
            conn_suf = conn_suf + "StarOffice.Service"
        conn = f'{prefix}"{conn_suf}"'
        args.append(conn)

        # print("_popen() Connection Args:", args)

        if shutdown == True:
            if self._is_nt:
                cmd_str = " ".join(args)
                self._soffice_process_shutdown = subprocess.Popen(
                    cmd_str, shell=True, env=self._envirnment
                )
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
            if self._is_nt:
                self._soffice_process = subprocess.Popen(
                    cmd_str, shell=True, env=self._envirnment
                )
            else:
                self._soffice_process = subprocess.Popen(
                    cmd_str,
                    env=self._envirnment,
                    preexec_fn=os.setsid,
                    shell=True,
                )
            # print("_popen() Process started")

    def _connect(self):

        identifier = f"socket,host={self._host},port={self._port}"
        conn_str = f"uno:{identifier};urp;StarOffice.ServiceManager"
        # ctx = resolver.resolve( "uno:socket, host = localhost, port= 2002; urp; StarOffice.ComponentContext" )

        for i in range(self._conn_try_count):
            try:
                localContext = uno.getComponentContext()
                resolver = localContext.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", localContext
                )
                smgr = resolver.resolve(conn_str)
                self._script_content = smgr.getPropertyValue("DefaultContext")
                break
            except NoConnectException as e:
                time.sleep(self._conn_try_time_sleep)
                if i == self._conn_try_count - 1:
                    raise e

class LoDirectStart(ConnectBase):
    def __init__(self, **kwargs):
        """
        Constructor

        Keyword Arguments:
            working_dir (str, optional): the working directory of LO. This is the path that profile will be copied into. Defaults to a tmp directory
            cache_path (str, optional): the path where the current profile resides to copy into the LO
                workding dir for current instance of LO
            start_soffice (bool, optional): If True soffice will be started a server that can be connected to. Default True
        """
        super().__init__(**kwargs)
   
    def connect(self):
        """
        Makes a connection to soffice

        Raises:
            NoConnectException: if unable to obtain a connection to soffice
        """
        self._script_content = uno.getComponentContext()
        