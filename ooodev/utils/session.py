# coding: utf-8
from __future__ import annotations, unicode_literals
import sys
from typing import TYPE_CHECKING, cast
import uno
import getpass, os, os.path
from ..meta.static_meta import StaticProperty, classproperty
from ..events.args.event_args import EventArgs
from ..events.lo_named_event import LoNamedEvent
from ..events.event_singleton import _Events
from . import lo as mLo
from ..exceptions import ex as mEx

from com.sun.star.util import XStringSubstitution
from ooo.dyn.uno.deployment_exception import DeploymentException

# com.sun.star.uno.DeploymentException

if TYPE_CHECKING:
    from com.sun.star.util import PathSubstitution
    from com.sun.star.uno import XComponentContext


class Session(metaclass=StaticProperty):
    """
    Session Class for handling user paths within the LibreOffice environment.

    See Also:
        - `Importing Python Modules <https://help.libreoffice.org/latest/lo/text/sbasic/python/python_import.html>`_
        - `Getting Session Information <https://help.libreoffice.org/latest/lo/text/sbasic/python/python_session.html>`_
    """

    # https://help.libreoffice.org/latest/lo/text/sbasic/python/python_import.html
    # https://help.libreoffice.org/latest/lo/text/sbasic/python/python_session.html

    class PathEnum:
        SHARE_PYTHON = 1
        SHARE_USER_PYTHON = 2

    @classproperty
    def path_sub(cls) -> PathSubstitution:
        try:
            return cls._path_substitution
        except AttributeError:
            try:
                # raises DeploymentException if not in a macro environment
                # in a macro envrionment there should be no dependency on Lo
                # this will allow for import shared python files before Lo.load_office is called.
                # if not in a macro then must get instace after Lo.load_office is called
                ctx = cast("XComponentContext", uno.getComponentContext())
                ps = ctx.getServiceManager().createInstanceWithContext("com.sun.star.util.PathSubstitution", ctx)
                cls._path_substitution = ps
                return cls._path_substitution
            except DeploymentException as e:
                # print(e)
                # there must be a connection to before calling session.
                if mLo.Lo.is_loaded is False:
                    raise mEx.LoNotLoadedError(
                        "Lo.load_office must be called before using session when not run as a macro"
                    )
                cls._path_substitution = mLo.Lo.create_instance_mcf(
                    XStringSubstitution, "com.sun.star.util.PathSubstitution"
                )
        return cls._path_substitution

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
            cls._share = os.path.normpath(inst.replace("program", "Share"))
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
    def register_path(cls, path: Session.PathEnum) -> None:
        """
        Registers a path into ``sys.path`` if it does not exist

        Args:
            path (PathEnum): Type of path to register.
        """
        if path == Session.PathEnum.SHARE_PYTHON:
            spath = cls.shared_py_scripts
        elif path == Session.PathEnum.SHARE_USER_PYTHON:
            spath = cls.user_py_scripts
        if not spath in sys.path:
            sys.path.insert(0, spath)


def _del_cache_attrs(source: object, e: EventArgs) -> None:
    # clears Lo Attributes that are dynamically created
    dattrs = ("_path_substitution", "_share", "_user_profile")
    for attr in dattrs:
        if hasattr(Session, attr):
            delattr(Session, attr)


_Events().on(LoNamedEvent.RESET, _del_cache_attrs)

__all__ = ("Session",)
