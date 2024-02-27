# coding: utf-8
from __future__ import annotations, unicode_literals
import sys
from typing import TYPE_CHECKING, cast
from enum import Enum
import getpass
import os
import os.path

import uno
from com.sun.star.util import XStringSubstitution
from ooo.dyn.uno.deployment_exception import DeploymentException

from ooodev.meta.static_meta import StaticProperty, classproperty
from ooodev.events.args.event_args import EventArgs
from ooodev.events.lo_named_event import LoNamedEvent
from ooodev.events.event_singleton import _Events
from ooodev.loader import lo as mLo
from ooodev.exceptions import ex as mEx


# com.sun.star.uno.DeploymentException

if TYPE_CHECKING:
    from com.sun.star.util import PathSubstitution  # service
    from com.sun.star.uno import XComponentContext

# pylint: disable=unused-argument


class PathKind(Enum):
    """Kind of path to register"""

    SHARE_PYTHON = 1
    SHARE_USER_PYTHON = 2


class Session(metaclass=StaticProperty):
    """
    Session Class for handling user paths within the LibreOffice environment.

    See Also:
        - `Importing Python Modules <https://help.libreoffice.org/latest/lo/text/sbasic/python/python_import.html>`_
        - `Getting Session Information <https://help.libreoffice.org/latest/lo/text/sbasic/python/python_session.html>`_
    """

    # https://help.libreoffice.org/latest/lo/text/sbasic/python/python_import.html
    # https://help.libreoffice.org/latest/lo/text/sbasic/python/python_session.html

    PathEnum = PathKind

    @classproperty
    def path_sub(cls) -> PathSubstitution:
        try:
            return cls._path_substitution  # type: ignore
        except AttributeError:
            try:
                # raises DeploymentException if not in a macro environment
                # in a macro environment there should be no dependency on Lo
                # this will allow for import shared python files before Lo.load_office is called.
                # if not in a macro then must get instance after Lo.load_office is called
                ctx = cast("XComponentContext", uno.getComponentContext())
                ps = ctx.getServiceManager().createInstanceWithContext("com.sun.star.util.PathSubstitution", ctx)
                cls._path_substitution = ps
                return cls._path_substitution  # type: ignore
            except DeploymentException as e:
                # print(e)
                # there must be a connection to before calling session.
                if mLo.Lo.is_loaded is False:
                    raise mEx.LoNotLoadedError(
                        "Lo.load_office must be called before using session when not run as a macro"
                    ) from e
                cls._path_substitution = mLo.Lo.create_instance_mcf(
                    XStringSubstitution, "com.sun.star.util.PathSubstitution", raise_err=True
                )
        return cls._path_substitution  # type: ignore

    @staticmethod
    def substitute(var_name: str):
        """
        Returns the current value of a variable.

        The method iterates through its internal variable list and tries to find the given variable.
        If the variable is unknown a com.sun.star.container.NoSuchElementException is thrown.

        Args:
            var_name (str): name to search for.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        # pylint: disable=no-member
        return Session.path_sub.getSubstituteVariableValue(var_name)

    @classproperty
    def share(cls) -> str:
        """
        Gets Program Share dir,
        such as ``C:\\Program Files\\LibreOffice\\share``
        """
        try:
            return cls._share
        except AttributeError:
            inst = uno.fileUrlToSystemPath(Session.substitute("$(prog)"))
            cls._share = os.path.normpath(inst.replace("program", "share"))
        return cls._share

    @classproperty
    def shared_scripts(cls) -> str:
        """
        Gets Program Share scripts dir,
        such as ``C:\\Program Files\\LibreOffice\\share\\Scripts``
        """
        return "".join([Session.share, os.sep, "Scripts"])

    @classproperty
    def shared_py_scripts(cls) -> str:
        """
        Gets Program Share python dir,
        such as ``C:\\Program Files\\LibreOffice\\share\\Scripts\\python``
        """
        # eg: C:\Program Files\LibreOffice\share\Scripts\python
        return "".join([Session.shared_scripts, os.sep, "python"])

    @classproperty  # alternative to '$(username)' variable
    def user_name(cls):
        """Get the username from the environment or password database.

        First try various environment variables, then the password
        database.  This works on Windows as long as USERNAME is set.
        """
        return getpass.getuser()

    @classproperty
    def user_profile(cls):
        """
        Gets path to user profile such as,
        ``C:\\Users\\user\\AppData\\Roaming\\LibreOffice\\4\\user``
        """
        try:
            return cls._user_profile
        except AttributeError:
            cls._user_profile = uno.fileUrlToSystemPath(Session.substitute("$(user)"))
        return cls._user_profile

    @classproperty
    def user_scripts(cls):
        """
        Gets path to user profile scripts such as,
        ``C:\\Users\\user\\AppData\\Roaming\\LibreOffice\\4\\user\\Scripts``
        """
        return "".join([cls.user_profile, os.sep, "Scripts"])

    @classproperty
    def user_py_scripts(cls):
        """
        Gets path to user profile python such as,
        ``C:\\Users\\user\\AppData\\Roaming\\LibreOffice\\4\\user\\Scripts\\python``
        """
        # eg: C:\Users\user\AppData\Roaming\LibreOffice\4\user\Scripts\python
        return "".join([cls.user_scripts, os.sep, "python"])

    @classmethod
    def register_path(cls, path: PathKind) -> None:
        """
        Registers a path into ``sys.path`` if it does not exist

        Args:
            path (PathKind): Type of path to register.
        """
        script_path = ""
        if path == PathKind.SHARE_PYTHON:
            script_path = cls.shared_py_scripts
        elif path == PathKind.SHARE_USER_PYTHON:
            script_path = cast(str, cls.user_py_scripts)
        if not script_path:
            return
        if script_path not in sys.path:
            sys.path.insert(0, script_path)


def _del_cache_attrs(source: object, e: EventArgs) -> None:
    # clears Lo Attributes that are dynamically created
    data_attrs = ("_path_substitution", "_share", "_user_profile")
    for attr in data_attrs:
        if hasattr(Session, attr):
            delattr(Session, attr)


_Events().on(LoNamedEvent.RESET, _del_cache_attrs)

__all__ = ("Session",)
