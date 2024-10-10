"""
The original getPackageName2PathMap method uses the thePackageManagerFactory service to get the packages
which is not deprecated and was causing some issus.

The PythonScriptProvider class is a wrapper around the original PythonScriptProvider class from the pythonscript module.
This is to use ExtensionManager to get the packages.
"""

from __future__ import annotations
import uno
from com.sun.star.uno import RuntimeException
import pythonscript  # type: ignore
from pythonscript import PythonScriptProvider as UnoPythonScriptProvider  # type: ignore
from pythonscript import DirBrowseNode as UnoDirBrowseNode  # type: ignore
from pythonscript import FileBrowseNode as UnoFileBrowseNode  # type: ignore
from pythonscript import PackageBrowseNode as UnoPackageBrowseNode  # type: ignore
from pythonscript import ScriptBrowseNode as UnoScriptBrowseNode  # type: ignore
from pythonscript import Package as UnoPackage  # type: ignore
from pythonscript import PythonScript as UnoPythonScript  # type: ignore
from pythonscript import ScriptContext as UnoScriptContext  # type: ignore
from pythonscript import ProviderContext as UnoProviderContext  # type: ignore

from ooodev.adapter.deployment.extension_manager_comp import ExtensionManagerComp


def getPackageName2PathMap(sfa, storageType):
    ret = {}

    emc = ExtensionManagerComp.from_lo()

    pythonscript.log.debug("pythonscript: getPackageName2PathMap start getDeployedPackages")
    packages = emc.get_deployed_extensions(
        repository=pythonscript.mapStorageType2PackageContext(storageType),
        abort_channel=emc.create_abort_channel(),
        cmd_env=pythonscript.CommandEnvironment(),
    )

    pythonscript.log.debug("pythonscript: getPackageName2PathMap end getDeployedPackages (" + str(len(packages)) + ")")

    for pkg in packages:
        pythonscript.log.debug("inspecting package " + pkg.Name + "(" + pkg.Identifier.Value + ")")
        transientPathElement = pythonscript.penultimateElement(pkg.URL)
        j = pythonscript.expandUri(pkg.URL)
        paths = pythonscript.getPathsFromPackage(j, sfa)
        if len(paths) > 0:
            # map package name to url, we need this later
            pythonscript.log.debug("adding Package " + transientPathElement + " " + str(paths))
            ret[pythonscript.lastElement(j)] = pythonscript.Package(paths, transientPathElement)
    return ret


class DirBrowseNode(UnoDirBrowseNode):
    """
    DirBrowseNode class.
    """

    pass


class FileBrowseNode(UnoFileBrowseNode):
    """
    FileBrowseNode class.
    """

    pass


class PackageBrowseNode(UnoPackageBrowseNode):
    """
    PackageBrowseNode class.
    """

    pass


class Package(UnoPackage):
    """
    Package class.
    """

    pass


class PythonScript(UnoPythonScript):
    """
    PythonScript class.
    """

    pass


class ScriptBrowseNode(UnoScriptBrowseNode):
    """
    ScriptBrowseNode class.
    """

    pass


class ScriptContext(UnoScriptContext):
    """
    ScriptContext class.
    """

    pass


class ProviderContext(UnoProviderContext):
    """
    ProviderContext class.
    """

    pass


class PythonScriptProvider(UnoPythonScriptProvider):
    """
    PythonScriptProvider class.

    This class is a wrapper around the original PythonScriptProvider class from the pythonscript module.
    It is wrapped so that it uses ExtensionManager to get the packages.
    The original PythonScriptProvider class uses the thePackageManagerFactory service to get the packages
    which is not deprecated and was causing some issus.

    Args:
        UnoPythonScriptProvider (_type_): _description_
    """

    def __init__(self, ctx, *args):
        if pythonscript.log.isDebugLevel():
            mystr = ""
            for i in args:
                if len(mystr) > 0:
                    mystr = mystr + ","
                mystr = mystr + str(i)
            pythonscript.log.debug("Entering PythonScriptProvider.ctor" + mystr)

        doc = None
        inv = None
        storageType = ""

        if isinstance(args[0], str):
            storageType = args[0]
            if storageType.startswith("vnd.sun.star.tdoc"):
                doc = pythonscript.getModelFromDocUrl(ctx, storageType)
        else:
            inv = args[0]
            try:
                doc = inv.ScriptContainer
                content = (
                    ctx.getServiceManager()
                    .createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
                    .createDocumentContent(doc)
                )
                storageType = content.getIdentifier().getContentIdentifier()
            except Exception as e:
                text = pythonscript.lastException2String()
                pythonscript.log.error(text)

        isPackage = storageType.endswith(":uno_packages")

        try:
            #            urlHelper = ctx.ServiceManager.createInstanceWithArgumentsAndContext(
            #                "com.sun.star.script.provider.ScriptURIHelper", (LANGUAGENAME, storageType), ctx)
            urlHelper = pythonscript.MyUriHelper(ctx, storageType)
            pythonscript.log.debug("got urlHelper " + str(urlHelper))

            rootUrl = pythonscript.expandUri(urlHelper.getRootStorageURI())
            pythonscript.log.debug(storageType + " transformed to " + rootUrl)

            ucbService = "com.sun.star.ucb.SimpleFileAccess"
            sfa = ctx.ServiceManager.createInstanceWithContext(ucbService, ctx)
            if not sfa:
                pythonscript.log.debug("PythonScriptProvider couldn't instantiate " + ucbService)
                raise RuntimeException("PythonScriptProvider couldn't instantiate " + ucbService, self)
            self.provCtx = ProviderContext(
                storageType, sfa, urlHelper, ScriptContext(uno.getComponentContext(), doc, inv)
            )
            if isPackage:
                mapPackageName2Path = getPackageName2PathMap(sfa, storageType)
                self.provCtx.setPackageAttributes(mapPackageName2Path, rootUrl)
                self.dirBrowseNode = PackageBrowseNode(self.provCtx, pythonscript.LANGUAGENAME, rootUrl)
            else:
                self.dirBrowseNode = DirBrowseNode(self.provCtx, pythonscript.LANGUAGENAME, rootUrl)

        except Exception as e:
            text = pythonscript.lastException2String()
            pythonscript.log.debug("PythonScriptProvider could not be instantiated because of : " + text)
            raise e
