from __future__ import annotations
from typing import Any, cast, overload, Tuple
import threading
import contextlib
from functools import lru_cache
import uno
from com.sun.star.script.provider import XScript
from com.sun.star.uno import XInterface

from ooodev.loader import lo as mLo


class MacroScript:
    """
    Class for call macro

    See Also:
        `See Scripting Framework <https://wiki.documentfoundation.org/Documentation/DevGuide/Scripting_Framework#Scripting_Framework_URI_Specification>`_

    .. versionadded:: 0.40.0
    """

    @classmethod
    def call_url(cls, url: str, *args: Any, in_thread: bool = False) -> Any:
        """
        Call macro by url

        Args:
            args (tuple): Arguments passed to script when invoking.
            url (str): The URL of the script to be invoked.
            in_thread (bool, optional): If execute in thread. Defaults to ``False``.

        Returns:
            Any: The return value of the script.
        """
        result = None
        if in_thread:
            t = threading.Thread(target=cls._call_url, args=(url, *args))
            t.start()
        else:
            result = cls._call_url(url, *args)
        return result

    @overload
    @classmethod
    def call(
        cls,
        *,
        name: str,
        library: str = "Standard",
        language: str = "Basic",
        location: str = "user",
        module: str = ".",
        in_thread: bool = False,
        args: Tuple[Any, ...] = (),
    ) -> Any:
        """
        Call any macro

        Args:
            name (str): Macro Name.
            library (str, optional): Macro Library. Defaults to ``Standard``.
            language (str, optional): Language ``Basic`` or ``Python``. Defaults to ``Basic``.
            location (str, optional): Location ``user``, ``application`` or ``document``. Defaults to "user".
            module (str, optional): Module portion. Only Applies if ``language` is not ``Basic`` or ``Python``. Defaults to ".".
            in_thread (bool, optional): If execute in thread. Defaults to ``False``.
            args (tuple, optional): Arguments passed to script when invoking. Defaults to ().
        """
        ...

    @overload
    @classmethod
    def call(cls, **kwargs: Any) -> Any: ...

    @classmethod
    def call(cls, **kwargs: Any) -> Any:
        """Call any macro

        Args:
            name (str): Macro Name.
            library (str, optional): Macro Library. Defaults to ``Standard``.
            language (str, optional): Language ``Basic`` or ``Python``. Defaults to ``Basic``.
            location (str, optional): Location ``user``, ``application`` or ``document``. Defaults to "user".
            module (str, optional): Module portion. Only Applies if ``language` is not ``Basic`` or ``Python``. Defaults to ".".
            in_thread (bool, optional): If execute in thread. Defaults to ``False``.
            args (tuple, optional): Arguments passed to script when invoking. Defaults to ().

        Returns:
            Any: The return value of the script.

        Example:
            The ``MacroScript.call()`` method:

            .. code-block:: python

                >>> def r_trim(input: str, remove: str = " ") -> str:
                ...     res = MacroScript.call(
                ...         name="RTrimStr",
                ...         library="Tools",
                ...         language="Basic",
                ...         module="Strings",
                ...         args=(input, remove),
                ...     )
                ...     return res
                >>> result = r_trim("hello ")
                >>> assert result == "hello"
        """
        in_thread = bool(kwargs.pop("in_thread", False))
        result = None
        if in_thread:
            t = threading.Thread(target=cls._call, args=(kwargs,))
            t.start()
        else:
            result = cls._call(**kwargs)
        return result

    @overload
    @classmethod
    def get_url_script(
        cls,
        *,
        name: str,
        library: str = "Standard",
        language: str = "Basic",
        location: str = "user",
        module: str = ".",
    ) -> str:
        """
        Get uno command or url for macro

        Args:
            name (str): Macro Name.
            library (str, optional): Macro Library. Defaults to ``Standard``.
            language (str, optional): Language ``Basic`` or ``Python``. Defaults to ``Basic``.
            location (str, optional): Location ``user`` or ``application``. Defaults to "user".
            module (str, optional): Module portion. Only Applies if ``language`` is not ``Basic`` or ``Python``. Defaults to ".".

        Returns:
            str: Macro Url such as ``vnd.sun.star.script:myLibrary.myModule.myMacro?language=Basic&location=application``.
        """
        ...

    @overload
    @classmethod
    def get_url_script(cls, **kwargs: Any) -> str:
        """
        Get uno command or url for macro

        Args:
            name (str): Macro Name.
            library (str, optional): Macro Library. Defaults to ``Standard``.
            language (str, optional): Language ``Basic`` or ``Python``. Defaults to ``Basic``.
            location (str, optional): Location ``user`` or ``application``. Defaults to "user".
            module (str, optional): Module portion. Only Applies if ``language` is not ``Basic`` or ``Python``. Defaults to ".".

        Returns:
            str: Macro Url such as ``vnd.sun.star.script:myLibrary.myModule.myMacro?language=Basic&location=application``.
        """
        ...

    @classmethod
    def get_url_script(cls, **kwargs: Any) -> str:
        """
        Get uno command or url for macro

        Args:
            name (str): Macro Name.
            library (str, optional): Macro Library. Defaults to ``Standard``.
            language (str, optional): Language ``Basic`` or ``Python``. Defaults to ``Basic``.
            location (str, optional): Location ``user``, ``application`` or ``document``. Defaults to "user".
            module (str, optional): Module portion. Only Applies if ``language` is not ``Basic`` or ``Python``. Defaults to ".".

        Returns:
            str: Macro Url such as ``vnd.sun.star.script:myLibrary.myModule.myMacro?language=Basic&location=application``.
        """
        name = kwargs["name"]
        library = kwargs.get("library", "Standard")
        language = kwargs.get("language", "Basic")
        location = kwargs.get("location", "user")
        module = kwargs.get("module", ".")
        lower_lang = language.lower()
        if lower_lang == "python":
            language = "Python"
            module = ".py$"
        elif lower_lang == "basic":
            language = "Basic"
            module = f".{module}."
            if location == "user":
                location = "application"

        url = "vnd.sun.star.script"
        url = f"{url}:{library}{module}{name}?language={language}&location={location}"
        return url

    @staticmethod
    @lru_cache(maxsize=30)
    def get_script(script_url: str) -> XScript:
        """
        Grab Basic script object before invocation.

        Args:
            script_url (str): The URL of the script to be invoked.

        Returns:
            XScript: Basic script object.
        """
        embedded = False
        with contextlib.suppress(ValueError):
            idx = script_url.index("location=")
            location = script_url[idx + 9 :]
            embedded = location.lower().startswith("document")
        if embedded:
            dk = cast(Any, mLo.Lo.desktop.component)
            script_pro = dk.CurrentComponent.getScriptProvider()
        else:
            master_provider = cast(
                Any, mLo.Lo.create_instance_mcf(XInterface, "com.sun.star.script.provider.MasterScriptProviderFactory")
            )
            script_pro = master_provider.createScriptProvider("")
        return script_pro.getScript(script_url)

    @classmethod
    def _call(cls, **kwargs) -> Any:
        url = cls.get_url_script(**kwargs)
        args = kwargs.get("args", ())
        return cls._call_url(url, *args)

    @classmethod
    def _call_url(cls, url: str, *args: Any) -> Any:
        script = MacroScript.get_script(url)
        return script.invoke(args, None, None)[0]  # type: ignore
