from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno
from ooodev.adapter.ucb.simple_file_access_comp import SimpleFileAccessComp
from ooodev.adapter.io.pipe_comp import PipeComp
from ooodev.adapter.io.text_output_stream_comp import TextOutputStreamComp
from ooodev.adapter.io.text_input_stream_comp import TextInputStreamComp
from ooodev.io.log.named_logger import NamedLogger

if TYPE_CHECKING:
    from com.sun.star.ucb import XSimpleFileAccess3


class Sfa:
    """
    Class that can access files using SimpleFileAccess service.

    This class can bridge from the document to the file system.

    Example:
        This example copies a file to a new directory in the document.
        When the document is saved the file will be saved in the document.

        .. code-block:: python

            from pathlib import Path
            from ooodev.io.sfa.sfa import Sfa
            # ... other code

            sfa = Sfa()
            root = f"vnd.sun.star.tdoc:/{doc.runtime_uid}/"
            new_dir = root + "new_dir"
            sfa.inst.create_folder(new_dir)
            myfile = Path("my_file.txt")
            s_file = new_dir + "/new_file.txt"
            sfa.inst.copy(source_url=myfile.as_uri(), dest_url=s_file)

    """

    DOC_PROTOCOL = "vnd.sun.star.tdoc"

    def __init__(self, sfa: XSimpleFileAccess3 | None = None):
        """
        Constructor

        Args:
            sfa (XSimpleFileAccess3 | None, optional): Simple File Access component. If omitted then new sfa is created. Defaults to None.
        """
        if sfa is None:
            self._sfa = SimpleFileAccessComp.from_lo()
        else:
            self._sfa = SimpleFileAccessComp(sfa)
        self._logger = NamedLogger(self.__class__.__name__)

    def read_text_file(self, uri: str) -> str:
        """
        Read content from file

        Args:
            uri (str): The name of the file such as ``vnd.sun.star.tdoc:/1/Scripts/python/MyFile.py`` or any other uri format supported by the SimpleFileAccess service.

        Raises:
            FileNotFoundError: If the file is not found.

        Returns:
            str: The content.
        """

        if not self._sfa.exists(uri):
            self._logger.error(f"read_text_file(): File Not Found. URI: {uri}")
            raise FileNotFoundError("File not found.")
        try:
            io = self._sfa.open_file_read(uri)
        except Exception:
            self._logger.error(f"read_text_file(): Error opening file. URI: {uri}")
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
        except Exception:
            raise

    def write_text_file(self, uri: str, content: str, mode: str = "w") -> None:
        """
        Write content to file.

        Args:
            uri (str): The name of the file such as ``vnd.sun.star.tdoc:/1/Scripts/python/MyFile.py`` or any other uri format supported by the SimpleFileAccess service.
            content (str): The content to write.
            mode: (str, optional): The mode to open the file. Defaults to "w".
                mode ``w`` will overwrite the file.
                mode ``a`` will append to the file.
                mode ``x`` will create a new file and write to it failing if the file already exists
        """
        if mode not in {"w", "a", "x"}:
            mode = "w"

        file_exist = self._sfa.exists(uri)
        if file_exist:
            if mode == "x":
                raise FileExistsError("File exist.")
            if mode == "w":
                self._sfa.kill(uri)
            else:
                # append
                contents = self.read_text_file(uri=uri)
                content = contents + "\n" + content
                self._sfa.kill(uri)
        is_doc = uri.startswith(self.DOC_PROTOCOL)
        try:
            if is_doc:
                io = PipeComp.from_lo().component
            else:
                io = self._sfa.open_file_write(uri)
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
                    self._sfa.write_file(uri, io)  # type: ignore
        except Exception:
            raise
        finally:
            if is_doc:
                io.closeInput()  # type: ignore
            else:
                io.closeOutput()

    def delete_file(self, uri: str) -> None:
        """
        Delete file.

        Args:
            uri (str): The name of the file such as ``vnd.sun.star.tdoc:/1/Scripts/python/MyFile.py`` or any other uri format supported by the SimpleFileAccess service.
        """
        if self._sfa.exists(uri):
            self._sfa.kill(uri)

    def exists(self, file_url: str) -> bool:
        """
        Checks if a file exists.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self._sfa.exists(file_url)

    def get_folder_contents(self, folder_url: str, include_folders: bool = True) -> Tuple[str, ...]:
        """
        Get the content of a folder (file uri's).

        Args:
            uri (str): The name of the folder such as ``vnd.sun.star.tdoc:/1/Scripts/python`` or any other uri format supported by the SimpleFileAccess service.
            include_folders (bool, optional): If ``True``, folders are included in the result. Defaults to ``True``.

        Returns:
            Tuple[str, ...]: Tuple of file uri's.
        """
        return self._sfa.get_folder_contents(folder_url=folder_url, include_folders=include_folders)

    @property
    def inst(self) -> SimpleFileAccessComp:
        """
        Get the SimpleFileAccess component.

        Returns:
            SimpleFileAccessComp: The SimpleFileAccess component.
        """
        return self._sfa
