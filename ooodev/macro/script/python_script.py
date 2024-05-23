from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
import unohelper
from ooodev.loader import lo as mLo
from ooodev.io.log.named_logger import NamedLogger
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.adapter.frame.transient_documents_document_content_factory_comp import (
    TransientDocumentsDocumentContentFactoryComp,
)
from ooodev.adapter.io.pipe_comp import PipeComp
from ooodev.adapter.io.text_output_stream_comp import TextOutputStreamComp
from ooodev.adapter.io.text_input_stream_comp import TextInputStreamComp
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.adapter.ucb.simple_file_access_comp import SimpleFileAccessComp

try:
    import pythonscript
except ImportError:
    import pythonloader

    pythonscript = None
    for url, module in pythonloader.g_loadedComponents.iteritems():
        if url.endswith("script-provider-for-python/pythonscript.py"):
            pythonscript = module
    if pythonscript is None:
        raise Exception("Unable to find the pythonscript module.")

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from ooodev.loader.inst.lo_inst import LoInst


class PythonScript(LoInstPropsPartial):
    """
    Python script access.

    This class provides access to Python scripts in the document, user, and shared script providers.

    This class can also be accessed via the ``python_script`` property of the various Documents such as ``Write``, ``Calc``, ``Draw``, ``Impress``.

    .. versionadded:: 0.43.0
    """

    FILE_EXT = ".py"
    DOC_PROTOCOL = "vnd.sun.star.tdoc"
    SCRIPT_PROTOCOL = "vnd.sun.star.script"

    def __init__(self, doc: XComponent | None = None, lo_inst: LoInst | None = None) -> None:
        """
        Constructor.

        Args:
            doc (XComponent | None, optional): Document Component. If omitted the defaults to current active document.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        self._allow_macro_execution = False
        self._user_sp = None
        self._share_sp = None
        self._document_sp = None
        self._ctx = lo_inst.get_context()
        self._doc = doc
        self._doc_uri = None

        self._logger = NamedLogger(self.__class__.__name__)

    def _join_url(self, base, name, name_encode=True):
        """Join name to base URL."""
        if name_encode:
            _name = name
        else:
            _name = unohelper.systemPathToFileUrl(name)
        if base.endswith("/"):
            return base + _name
        return base + "/" + _name

    def _get_document_uri(self, doc: XComponent):
        """Get document transient URI."""
        transient_doc = TransientDocumentsDocumentContentFactoryComp.from_lo(self.lo_inst)

        try:
            content = transient_doc.create_document_content(doc)  # type: ignore
            id = content.get_identifier()
            return id.getContentIdentifier()
        except Exception:
            pass

    def _get_doc_uri(self):
        """Returns internal uri for the current document."""
        if self._doc_uri:
            return self._doc_uri
        if self._doc is None:
            self._doc = cast(Any, mLo.Lo.desktop.get_current_component())
        try:
            # script container is another document
            # (typically for database sub-documents)
            doc = self._doc.ScriptContainer  # type: ignore
        except AttributeError:
            # some sub-documents don't support XScriptInvocationContext
            try:
                if self._doc.Parent:  # type: ignore
                    doc = self._doc.Parent  # type: ignore
            except AttributeError:
                pass
        uri = self._get_document_uri(doc)
        if uri:
            self._allow_macro_execution = doc.AllowMacroExecution
            self._doc_uri = uri.rstrip("/")
            return self._doc_uri
        return ""

    def _get_user_script_provider(self):
        """Get the user script provider."""
        if self._user_sp is None:
            PythonScriptProvider = pythonscript.PythonScriptProvider
            self._user_sp = PythonScriptProvider(self._ctx, "user")
            try:
                self._user_sp.uno_packages_sp = PythonScriptProvider(self._ctx, "user:uno_packages")
            except Exception:
                self._user_sp.uno_packages_sp = None
        return self._user_sp

    def _get_shared_script_provider(self):
        """Get the share script provider."""
        if self._share_sp is None:
            PythonScriptProvider = pythonscript.PythonScriptProvider
            self._share_sp = PythonScriptProvider(self._ctx, "share")
            try:
                self._share_sp.uno_packages_sp = PythonScriptProvider(self._ctx, "share:uno_packages")
            except Exception:
                self._share_sp.uno_packages_sp = None
        return self._share_sp

    def _get_document_script_provider(self):
        """Get the document script provider."""
        if self._document_sp is None:
            doc = self._get_doc_uri()
            PythonScriptProvider = pythonscript.PythonScriptProvider
            self._document_sp = PythonScriptProvider(self._ctx, doc)
            # don't create a default PythonScriptProvider instance due to bug https://bugs.documentfoundation.org/show_bug.cgi?id=105609
            # self.document_sp = default_sp.setdefault(self.doc, PythonScriptProvider(self.ctx, self.doc))
            self._document_sp.uno_packages_sp = None
        return self._document_sp

    def file_exist(self, filename: str, node: Any = None) -> bool:
        """
        Check if file exists under supplied node.

        Args:
            filename (str): The name of the file such as ``my_script`` or ``my_script.py``.
            node (Any, optional): The node to check.
                If omitted then the content will be checked in the document script provider.

        Returns:
            bool: True if file exists.
        """
        if node is None:
            node = self.document_script_provider.dirBrowseNode
        if isinstance(node, pythonscript.PythonScriptProvider):
            node = node.dirBrowseNode
        elif not isinstance(node, pythonscript.DirBrowseNode):
            self._logger.debug("file_exist(): Invalid node type.")
            return False
        name = filename
        if name and name.strip():
            uri = self._join_url(node.rootUrl, name)
            if not uri.endswith(self.FILE_EXT):
                uri += self.FILE_EXT
            sfa = node.provCtx.sfa
            return sfa.exists(uri)
        return False

    def delete_file(self, filename: str, node: Any = None) -> bool:
        """
        Delete file under supplied node.

        Args:
            filename (str): The name of the file such as ``my_script`` or ``my_script.py``.
            node (Any, optional): The node to delete from.
                If omitted then the content will be deleted from the document script provider.

        Returns:
            bool: True if file was deleted.
        """
        if node is None:
            node = self.document_script_provider.dirBrowseNode
        if isinstance(node, pythonscript.PythonScriptProvider):
            node = node.dirBrowseNode
        elif not isinstance(node, pythonscript.DirBrowseNode):
            self._logger.debug("delete_file(): Invalid node type.")
            return False
        name = filename
        if name and name.strip():
            uri = self._join_url(node.rootUrl, name)
            if not uri.endswith(self.FILE_EXT):
                uri += self.FILE_EXT
            sfa = node.provCtx.sfa
            if sfa.exists(uri):
                sfa.kill(uri)
                return True
        return False

    def _get_sfa(self) -> SimpleFileAccessComp:
        """Get SimpleFileAccessComp instance."""
        sfa = SimpleFileAccessComp.from_lo(self.lo_inst)
        return sfa

    def _create_string_node(self, node: str) -> DotDict:
        """Get string node."""
        sfa = self._get_sfa()
        s = f"{self._get_doc_uri()}/Scripts/python/{node.lstrip('/')}"
        return DotDict(rootUrl=s, provCtx=DotDict(sfa=sfa.component))

    # def _create_tempfile(self, node):
    #     """ Copy embedded script to temporary folder."""
    #     try:
    #         self.adddoceventlistener()
    #         if not self.tempdir:
    #             TP = self.smgr.createInstanceWithContext(
    #                         "com.sun.star.io.TempFile", self.ctx)
    #             self.tempdir = base_url(TP.Uri)
    #         path = node.uri.replace(self.DOC_PROTOCOL + ':/', '')
    #         filepath = join_url(self.tempdir, path)
    #         dirpath = '/'.join(filepath.split('/')[:-1])
    #         sfa = node.provCtx.sfa
    #         sfa.createFolder(dirpath)
    #         sfa.copy(node.uri, filepath)
    #         tempfiles[node.uri] = filepath
    #         return filepath
    #     except Exception as e:
    #         raise ErrorAsMessage("_create_tempfile\n\n"+str(e))

    def write_file(self, filename: str, content: str, node: Any = None, mode: str = "w") -> None:
        """
        Write content to file under supplied node.

        Args:
            filename (str): The name of the file such as ``my_script`` or ``my_script.py``.
            content (str): The content to write.
            node (Any, optional): The node to write to.
                If omitted then the content will be written to the document script provider.
            mode: (str, optional): The mode to open the file. Defaults to "w".
                mode ``w`` will overwrite the file.
                mode ``a`` will append to the file.
                mode ``x`` will create a new file and write to it failing if the file already exists
        """
        if mode not in {"w", "a", "x"}:
            mode = "w"

        if node is None:
            node = self.document_script_provider.dirBrowseNode
        if isinstance(node, pythonscript.PythonScriptProvider):
            node = node.dirBrowseNode
        elif isinstance(node, str):
            node = self._create_string_node(node)
        elif isinstance(node, DotDict):
            pass
        elif not isinstance(node, pythonscript.DirBrowseNode):
            self._logger.debug("write_file(): Invalid node type.")
            return
        name = filename
        if name and name.strip():
            # node.rootUrl = 'vnd.sun.star.tdoc:/1/Scripts/python'
            uri = self._join_url(node.rootUrl, name)
            if not uri.endswith(self.FILE_EXT):
                uri += self.FILE_EXT
            sfa = node.provCtx.sfa
            file_exist = sfa.exists(uri)
            if file_exist:
                if mode == "x":
                    raise FileExistsError("File exist.")
                if mode == "w":
                    sfa.kill(uri)
                else:
                    # append
                    contents = self.read_file(filename, node)
                    content = contents + "\n" + content
                    sfa.kill(uri)
            is_doc = uri.startswith(self.DOC_PROTOCOL)
            try:
                if is_doc:
                    io = PipeComp.from_lo().component
                else:
                    io = sfa.openFileWrite(uri)
            except Exception:
                raise
            try:
                if content or is_doc:
                    text_out = TextOutputStreamComp.from_lo()
                    text_out.set_output_stream(io)
                    text_out.set_encoding("UTF-8")
                    text_out.write_string(content)
                    if is_doc:
                        text_out.close_output()
                        sfa.writeFile(uri, io)
            except Exception:
                raise
            finally:
                if is_doc:
                    io.closeInput()
                else:
                    io.closeOutput()

    def read_file(self, filename: str, node: Any = None) -> str:
        """
        Read content from file under supplied node.

        Args:
            filename (str): The name of the file such as ``my_script`` or ``my_script.py``.
            node (Any, optional): The node to read from.
                If omitted then the content will be read from the document script provider.

        Returns:
            str: The content.
        """
        if node is None:
            node = self.document_script_provider.dirBrowseNode
        if isinstance(node, pythonscript.PythonScriptProvider):
            node = node.dirBrowseNode
        elif isinstance(node, str):
            node = self._create_string_node(node)
        elif isinstance(node, DotDict):
            pass
        elif not isinstance(node, pythonscript.DirBrowseNode):
            self._logger.debug("read_file(): Invalid node type.")
            return ""
        name = filename
        if name and name.strip():
            uri = self._join_url(node.rootUrl, name)
            if not uri.endswith(self.FILE_EXT):
                uri += self.FILE_EXT
            sfa = node.provCtx.sfa
            if not sfa.exists(uri):
                self._logger.error(f"read_file(): File Not Found. URI: {uri}")
                raise FileNotFoundError("File not found.")
            try:
                io = sfa.openFileRead(uri)
            except Exception:
                self._logger.error(f"read_file(): Error opening file. URI: {uri}")
                raise
            try:
                txt_stream = TextInputStreamComp.from_lo()
                txt_stream.set_input_stream(io)
                txt_stream.set_encoding("UTF-8")
                lines = []
                while not txt_stream.is_eof():
                    lines.append(txt_stream.read_string(False))
                txt_stream.close_input()
                result = "".join(lines)
                return result
                # str_lst = TextStream.get_str_list_from_text_input_stream(txt_stream, "\n", True)
                # content = str(str_lst)
                # return content
            except Exception:
                raise
            # finally:
            #     io.closeOutput()
        return ""

    def is_valid_python(self, code: str) -> bool:
        """
        Check if a string is valid Python code.

        This method can be used before write_file() to check if the code is valid.

        Args:
            code (str): The string to check.

        Returns:
            bool: True if the string is valid Python code, False otherwise.
        """
        try:
            compile(code, "<string>", "exec")
            return True
        except SyntaxError:
            self._logger.error("is_valid_python(): Syntax error.", exc_info=True)
            return False

    def test_compile_python(self, code: str) -> None:
        """
        Test compile Python code.

        Args:
            code (str): The code to compile.

        Raises:
            SyntaxError: If there is a syntax error.
        """
        try:
            compile(code, "<string>", "exec")
        except SyntaxError:
            self._logger.error("test_compile_python(): Syntax error.", exc_info=True)
            raise

    @property
    def document_script_provider(self):
        """Get the document script provider."""
        return self._get_document_script_provider()

    @property
    def shared_script_provider(self):
        """Get the shared script provider."""
        return self._get_shared_script_provider()

    @property
    def user_script_provider(self):
        """Get the user script provider."""
        return self._get_user_script_provider()
