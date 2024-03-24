from __future__ import annotations
from typing import cast, Any
from functools import lru_cache
import uno
from com.sun.star.script.provider import XScript
from com.sun.star.uno import XInterface

from ooodev.loader import lo as mLo


class Basic:
    """
    Class for managing Basic script.

    .. versionadded:: 0.38.0
    """

    @staticmethod
    @lru_cache(maxsize=30)
    def get_basic_script(macro="Main", module="Module1", library="Standard", embedded=False) -> XScript:
        """
        Grab Basic script object before invocation.

        Args:
            macro (str, optional): Macro name. Defaults to "Main".
            module (str, optional): Module name. Defaults to "Module1".
            library (str, optional): Library name. Defaults to "Standard".
            embedded (bool, optional): If ``True``, the script is embedded in the document. Defaults to ``False``.

        Returns:
            XScript: Basic script object.

        Example:
            The ``invoke`` method args:

            - the first lists the arguments of the called routine
            - the second identifies modified arguments
            - the third stores the return values

            .. code-block:: python

                >>> def r_trim(input: str, remove: str = " ") -> str:
                ...     script = Basic.get_basic_script(macro="RTrimStr", module="Strings", library="Tools", embedded=False)
                ...     res = script.invoke((input, remove), (), ())
                ...     return res[0]
                >>> result = r_trim("hello ")
                >>> assert result == "hello"

        See Also:
            LibreOffice - `Calling Basic Macros from Python <https://help.libreoffice.org/latest/en-US/text/sbasic/python/python_2_basic.html>`__
        """
        dk = cast(Any, mLo.Lo.desktop.component)
        if embedded:
            script_pro = dk.CurrentComponent.getScriptProvider()
            location = "document"
        else:
            master_provider = cast(
                Any, mLo.Lo.create_instance_mcf(XInterface, "com.sun.star.script.provider.MasterScriptProviderFactory")
            )
            script_pro = master_provider.createScriptProvider("")
            location = "application"
        script_name = f"vnd.sun.star.script:{library}.{module}.{macro}?language=Basic&location={location}"
        return script_pro.getScript(script_name)
