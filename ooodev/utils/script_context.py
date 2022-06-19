# coding: utf-8
# Origin class https://gitlab.com/LibreOfficiant/ide_utils/-/blob/master/IDE_utils.py
# Original Author: LibreOfficiant: https://gitlab.com/LibreOfficiant
from __future__ import annotations
from typing import TYPE_CHECKING
import uno
import unohelper
import os
from com.sun.star.script.provider import XScriptContext
if TYPE_CHECKING:
    from com.sun.star.frame import XModel
    from com.sun.star.uno import XComponentContext
    from com.sun.star.frame import XDesktop

class ScriptContext(unohelper.Base, XScriptContext):
    """ Substitute (Libre|Open)Office XSCRIPTCONTEXT built-in

    Can be used in IDEs such as Anaconda, Geany, KDevelop, PyCharm, VS Code..
    in order to run/debug Python macros.

    Implements: com.sun.star.script.provider.XScriptContext
    """
    def __init__(self, ctx, desktop, doc):
        self.ctx = ctx
        self.desktop = desktop
        self.doc = doc
    def getComponentContext(self) -> XComponentContext:
        return self.ctx
    def getDesktop(self) -> XDesktop:
        # return self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        return self.desktop
    def getDocument(self) -> XModel:
        # return self.getDesktop().getCurrentComponent()
        return self.doc
    def getInvocationContext(self):
        raise os.NotImplementedError

