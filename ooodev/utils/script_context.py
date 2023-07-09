# coding: utf-8
# Origin class https://gitlab.com/LibreOfficiant/ide_utils/-/blob/master/IDE_utils.py
# Original Author: LibreOfficiant: https://gitlab.com/LibreOfficiant
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
import os

_ON_RTD = os.environ.get("READTHEDOCS", None) == "True"
if _ON_RTD:
    from ..mock import unohelper
else:
    import unohelper

from com.sun.star.script.provider import XScriptContext

if TYPE_CHECKING:
    from com.sun.star.frame import XDesktop
    from com.sun.star.frame import XModel
    from com.sun.star.uno import XComponentContext
    from com.sun.star.lang import XComponent
    from com.sun.star.document import XScriptInvocationContext
else:
    XDesktop = object
    XModel = object
    XComponentContext = object
    XComponent = object
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
        doc: XComponent | None = None,
        inv: XScriptInvocationContext | None = None,
        dynamic: bool = True,
    ):
        self._uno_desktop_type = uno.getTypeByName("com.sun.star.frame.XDesktop")
        self._uno_model_type = uno.getTypeByName("com.sun.star.frame.XModel")
        self.ctx = ctx
        self.inv = inv
        self._dynamic = dynamic
        self._desktop = None
        if doc is None:
            self.doc = None
        else:
            self.doc = cast(XModel, doc.queryInterface(self._uno_model_type))

    def getComponentContext(self) -> XComponentContext:
        return self.ctx

    def getDesktop(self) -> XDesktop:
        # return self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        if not self._dynamic and self._desktop is not None:
            return self._desktop
        interface = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        self._desktop = cast(XDesktop, interface.queryInterface(self._uno_desktop_type))
        return self._desktop

    def getDocument(self) -> XModel:
        # return self.getDesktop().getCurrentComponent()
        if not self._dynamic and self.doc is not None:
            return self.doc
        component = self.getDesktop().getCurrentComponent()
        return cast(XModel, component.queryInterface(self._uno_model_type))

    def getInvocationContext(self) -> XScriptInvocationContext:
        if self.inv:
            return self.inv
        raise NotImplementedError
