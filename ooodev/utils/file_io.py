# coding: utf-8
# Python conversion of FileIO.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
from __future__ import annotations
import os
import sys
import tempfile
import datetime
import glob
import zipfile
from pathlib import Path
from urllib.parse import urlparse
from typing import Iterable, List, TYPE_CHECKING
import uno
from com.sun.star.uno import Exception as UnoException

if TYPE_CHECKING:
    from com.sun.star.packages.zip import XZipFileAccess
    from com.sun.star.container import XNameAccess
    from com.sun.star.io import XInputStream
    from com.sun.star.io import XTextInputStream
    from com.sun.star.io import XActiveDataSink

from . import lo as Util

if sys.version_info >= (3, 10):
    from typing import Union
else:
    from typing_extensions import Union

# Lo = m_lo.Lo

_UTIL_PATH = str(Path(__file__).parent)


class FileIO:
    @staticmethod
    def get_utils_folder() -> str:
        """
        Gets path to utils folder

        Returns:
            str: folder path as str
        """
        return _UTIL_PATH

    @staticmethod
    def get_absolute_path(fnm: str) -> str:
        """
        Gets Absolute path

        Args:
            fnm (str): path as string

        Returns:
            str: absolute path
        """
        return os.path.abspath(fnm)

    @staticmethod
    def url_to_path(url: str) -> str | None:
        """
        Converts url to path

        Args:
            url (str): url to convert

        Returns:
            str | None: path as str if conversion is successful; Otherwise, None
        """
        try:
            p = urlparse(url)
            final_path = os.path.abspath(os.path.join(p.netloc, p.path))
            return final_path
        except Exception as e:
            print(f"Could not parse '{url}'")
        return None

    @staticmethod
    def is_openable(fnm: str) -> bool:
        """
        Gets if a file can be opened

        Args:
            fnm (str): file path

        Returns:
            bool: True if file can be opened; Otherwise, False
        """
        try:
            p = Path(fnm)
            if not p.exists():
                print(f"'{fnm}' does not exist")
                return False
            if not p.is_file():
                print(f"'{fnm}' does is not a file")
                return False
            if not os.access(fnm, os.R_OK):
                print(f"'{fnm}' is not readable")
                return False
            return True
        except Exception as e:
            print(f"File is not openable: {e}")
        return False

    @staticmethod
    def fnm_to_url(fnm: str) -> str | None:
        """
        Converts file path to url

        Args:
            fnm (str): file path

        Returns:
            str | None: Converted path if conversion is successful; Otherwise None.
        """
        try:
            p = Path(fnm)
            return p.as_uri()
        except Exception as e:
            print("Unable to convert '{fnm}'")
        return None

    @classmethod
    def uri_to_path(cls, uri_fnm: str) -> str:
        return cls.url_to_path(url=uri_fnm)

    @staticmethod
    def make_directory(dir: str | Path) -> None:
        """
        Creates path and subpaths not existing.

        Args:
            dest_dir (str | Path): PathLike object
        """
        # Python ≥ 3.5
        if isinstance(dir, Path):
            dir.mkdir(parents=True, exist_ok=True)
        else:
            Path(dir).mkdir(parents=True, exist_ok=True)

    @staticmethod
    def get_file_names(dir: str) -> List[str]:
        """
        Gets a list of filenames in a folder

        Args:
            dir (str): Folder path

        Returns:
            List[str]: List of files.
        """
        # pattern .* includes hidden files whereas * does not.
        files = glob.glob(f"{dir}/.*", recursive=False)
        return files

    @staticmethod
    def get_fnm(path: str) -> str:
        """
        Gets last part of a file or dir such as myfile.txt

        Args:
            path (str): file path

        Returns:
            str: file name portion
        """
        if path == "":
            print("path is an empty string")
            return ""
        try:
            p = Path(path)
            return p.name
        except Exception as e:
            print(f"Unable to get name for '{path}'")
            print(f"    {e}")
        return ""

    @staticmethod
    def create_temp_file(im_format: str) -> str | None:
        """
        Creates a temporary file

        Args:
            im_format (str): File suffix such as txt or cfg

        Returns:
            str | None: Path to temp file.
        """
        try:
            tmp = tempfile.NamedTemporaryFile(prefix="loTemp", sufix=f".{im_format}", delete=True)
            return tmp.name
        except Exception as e:
            print("Could not create temp file")
        return None

    @staticmethod
    def delete_file(fnm: str) -> None:
        os.remove(fnm)
        if os.path.exists(fnm):
            print(f"'{fnm}' could not be deleted")
        else:
            print(f"'{fnm}' deleted")

    @classmethod
    def delete_files(cls, db_fnms: Iterable[str]) -> None:
        print()
        for s in db_fnms:
            cls.delete_file(s)

    @staticmethod
    def save_string(fnm: str, s: str) -> None:
        if s is None:
            print(f"No data to save in '{fnm}'")
        try:
            with open(fnm, "w") as file:
                file.write(s)
            print(f"Saved string to file: {fnm}")
        except Exception as e:
            print(f"Could not save string to file: {fnm}")

    @staticmethod
    def save_bytes(fnm: str, b: bytes) -> None:
        if b is None:
            print(f"No data to save in '{fnm}'")
        try:
            with open(fnm, "b") as file:
                file.write(b)
            print(f"Saved bytes to file: {fnm}")
        except Exception as e:
            print(f"Could not save bytes to file: {fnm}")

    @staticmethod
    def save_array(fnm: str, arr: List[list]) -> None:
        """
        Saves a 2d array to a file as tab delimited data.

        Args:
            fnm (str): file to save data to
            arr (List[list]): 2d array of data.
        """
        if arr is None:
            print("No data to save in '{fnm}'")
            return

        try:
            with open(fnm, "w") as file:
                if num_rows == 0:
                    print("No data to save in '{fnm}'")
                    return
                num_rows = len(arr)
                for j in range(num_rows):
                    line = "\t".join([str(v) for v in arr[j]])
                    file.write(line)
                    file.write("\n")
            print(f"Save array to file: {fnm}")
        except Exception as e:
            print(f"Could not save array to file: {fnm}")
            print(f"    {e}")

    @staticmethod
    def append_to(fnm: str, msg: str) -> None:
        try:
            with open(fnm, "a") as file:
                file.write(msg)
                file.write("\n")
        except Exception as e:
            print(f"unable to append to '{fnm}'")
            print(f"    {e}")

    # ----------------------- zip access ---------------------------------------
    @classmethod
    def zip_access(cls, fnm: str) -> XZipFileAccess:
        return Util.Lo.create_instance_mcf(
            XZipFileAccess, "com.sun.star.packages.zip.ZipFileAccess", (cls.fnm_to_url(fnm),)
        )

    @classmethod
    def zip_list_uno(cls, fnm: str) -> None:
        """Use zip_list method"""
        # replaced by more detailed Java version; see below
        zfa: XNameAccess = cls.zip_access(fnm)
        names = zfa.getElementNames()
        print(f"\nZippendContents of '{fnm}'")
        Util.Lo.print_names(names, 1)

    @staticmethod
    def unzip_file(zfa: XZipFileAccess, fnm: str) -> None:
        raise NotImplementedError

    @staticmethod
    def read_lines(in_stream: XInputStream) -> List[str] | None:
        lines = []
        try:
            tis = Util.Lo.create_instance_mcf(XTextInputStream, "com.sun.star.io.TextInputStream")
            sink = Util.Lo.qi(XActiveDataSink, tis)
            sink.setInputStream(in_stream)

            while tis.isEOF() is False:
                lines.append(tis.readLine())
            tis.closeInput()
        except Exception as e:
            print(e)
        if len(lines) == 0:
            return None
        return lines

    @classmethod
    def get_mime_type(cls, zfa: XZipFileAccess) -> str | None:
        try:
            in_stream: XInputStream = zfa.getStreamByPattern("mimetype")
            lines = cls.read_lines(in_stream)
            if lines is not None:
                return lines[0].strip()
        except UnoException as e:
            print(e)
        print("No mimetype found")
        return None

    # ----------------- switch to Python's zip APIs ------------

    @staticmethod
    def zip_list(fnm: str) -> None:
        try:
            with zipfile.ZipFile(fnm, "r") as zip:
                for info in zip.getinfo():
                    print(info.filename)
                    print("\tModified:\t" + str(datetime.datetime(*info.date_time)))
                    print("\tSystem:\t\t" + str(info.create_system) + "(0 = Windows, 3 = Unix)")
                    print("\tZIP version:\t" + str(info.create_version))
                    print("\tCompressed:\t" + str(info.compress_size) + " bytes")
                    print("\tUncompressed:\t" + str(info.file_size) + " bytes")
            print()
        except Exception as e:
            print(e)
