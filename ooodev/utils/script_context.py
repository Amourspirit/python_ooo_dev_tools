# coding: utf-8
# Origin class https://gitlab.com/LibreOfficiant/ide_utils/-/blob/master/IDE_utils.py
# Original Author: LibreOfficiant: https://gitlab.com/LibreOfficiant
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
import os
from com.sun.star.script.provider import XScriptContext
from ooodev.mock import mock_g

if mock_g.DOCS_BUILDING:
    from ooodev.mock import unohelper
else:
    import unohelper


if TYPE_CHECKING:
    from com.sun.star.frame import XDesktop
    from com.sun.star.frame import XModel
    from com.sun.star.uno import XComponentContext
    from com.sun.star.document import XScriptInvocationContext
else:
    XDesktop = object
    XModel = object
    XComponentContext = object
    XScriptInvocationContext = object


class ScriptContext(unohelper.Base, XScriptContext):  # type: ignore
    """Substitute (Libre|Open)Office XSCRIPTCONTEXT built-in

    Can be used in IDEs such as Anaconda, Geany, KDevelop, PyCharm, VS Code..
    in order to run/debug Python macros.

    Implements: com.sun.star.script.provider.XScriptContext
    """

    def __init__(
        self,
        ctx: XComponentContext,
        **kwargs,
    ):
        self._uno_desktop_type = uno.getTypeByName("com.sun.star.frame.XDesktop")
        self._uno_model_type = uno.getTypeByName("com.sun.star.frame.XModel")
        self._ctx = ctx
        self._inv: Any = kwargs.get("inv", None)
        self._desktop = None

    def getComponentContext(self) -> XComponentContext:
        return self._ctx

    def getDesktop(self) -> XDesktop:
        if self._desktop:
            return self._desktop
        interface = self._ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self._ctx)
        self._desktop = cast(XDesktop, interface.queryInterface(self._uno_desktop_type))
        return self._desktop

    def getDocument(self) -> XModel:
        component = self.getDesktop().getCurrentComponent()
        return cast(XModel, component.queryInterface(self._uno_model_type))

    def getInvocationContext(self) -> XScriptInvocationContext:
        if self._inv:
            return self._inv
        raise NotImplementedError
