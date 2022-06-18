# coding: utf-8
# Origin class https://gitlab.com/LibreOfficiant/ide_utils/-/blob/master/IDE_utils.py
# Original Author: LibreOfficiant: https://gitlab.com/LibreOfficiant
from __future__ import annotations
import uno
import unohelper
import os
from com.sun.star.script.provider import XScriptContext

class ScriptContext(unohelper.Base, XScriptContext):
    """ Substitute (Libre|Open)Office XSCRIPTCONTEXT built-in

    Can be used in IDEs such as Anaconda, Geany, KDevelop, PyCharm..
    in order to run/debug Python macros.

    Implements: com.sun.star.script.provider.XScriptContext

    Usage:

    ctx = connect(pipe='RichardMStalman')
    XSCRIPTCONTEXT = ScriptContext(ctx)

    ctx = connect(host='localhost',port=1515)
    XSCRIPTCONTEXT = ScriptContext(ctx)

    see also: `Runner`
    """
    '''
    cf. <OFFICEPATH>/program/pythonscript.py
    cf. https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=53748
    https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.xhtml?search=8100
    '''
    def __init__(self, ctx):
        self.ctx = ctx
    def getComponentContext(self):
        return self.ctx
    def getDesktop(self):
        return self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
    def getDocument(self):
        return self.getDesktop().getCurrentComponent()
    def getInvocationContext(self):
        raise os.NotImplementedError

