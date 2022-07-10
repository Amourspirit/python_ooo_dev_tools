# coding: utf-8
# Python conversion of FileIO.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
# region Imports
from __future__ import annotations
import os
import tempfile
import datetime
import glob
import zipfile
from pathlib import Path
from urllib.parse import urlparse
from typing import List, TYPE_CHECKING
import uno
from com.sun.star.io import XActiveDataSink
from com.sun.star.io import XTextInputStream
from com.sun.star.packages.zip import XZipFileAccess
from com.sun.star.uno import Exception as UnoException

if TYPE_CHECKING:
    from com.sun.star.container import XNameAccess
    from com.sun.star.io import XInputStream

from . import lo as mLo

from .type_var import PathOrStr

# if sys.version_info >= (3, 10):
#     from typing import Union
# else:
#     from typing_extensions import Union
# endregion imports

_UTIL_PATH = str(Path(__file__).parent)


class FileIO:
    
    # region ------------- file path methods ---------------------------
    @staticmethod
    def get_utils_folder() -> str:
        """
        Gets path to utils folder

        Returns:
            str: folder path as str
        """
        return _UTIL_PATH

    @staticmethod
    def get_absolute_path(fnm: PathOrStr) -> Path:
        """
        Gets Absolute path

        Args:
            fnm (PathOrStr): path as string

        Returns:
            Path: absolute path
        """
        p = Path(fnm)
        if p.is_absolute():
            return p
        return p.resolve()

    @classmethod
    def url_to_path(cls, url: str) -> Path:
        """
        Converts url to path

        Args:
            url (str): url to convert
        
        Raises:
            Exception: If unable to parse url.

        Returns:
            Path: path as string
        """
        try:
            p = urlparse(url)
            final_path = cls.get_absolute_path(os.path.join(p.netloc, p.path))
            return final_path
        except Exception as e:
            raise Exception(f"Could not parse '{url}'")


    @classmethod
    def fnm_to_url(cls, fnm: PathOrStr) -> str:
        """
        Converts file path to url

        Args:
            fnm (PathOrStr): file path

        Raises:
            Exception: If unalbe to get url form fnm.

        Returns:
            str: Converted path if conversion is successful; Otherwise None.
        """
        try:
            p =  cls.get_absolute_path(fnm)
            return p.as_uri()
        except Exception as e:
            raise Exception("Unable to convert '{fnm}'") from e

    @classmethod
    def uri_to_path(cls, uri_fnm: str) -> Path:
        """
        Convets uri file to path.

        Args:
            uri_fnm (str): Uri to convert

        Returns:
            Path: Converted uri as path.
        """
        return cls.url_to_path(url=uri_fnm)

    @staticmethod
    def get_file_names(dir: PathOrStr) -> List[str]:
        """
        Gets a list of filenames in a folder

        Args:
            dir (PathOrStr): Folder path

        Returns:
            List[str]: List of files.
        """
        # pattern .* includes hidden files whereas * does not.
        files = glob.glob(f"{dir}/.*", recursive=False)
        return files

    @staticmethod
    def get_fnm(path: PathOrStr) -> str:
        """
        Gets last part of a file or dir such as myfile.txt

        Args:
            path (PathOrStr): file path

        Returns:
            str: file name portion
        """
        if path == "":
            mLo.Lo.print("path is an empty string")
            return ""
        try:
            p = Path(path)
            return p.name
        except Exception as e:
            mLo.Lo.print(f"Unable to get name for '{path}'")
            mLo.Lo.print(f"    {e}")
        return ""
    
    # endregion ---------- file path methods ---------------------------

    # region ------------- file creation / deletion --------------------

    @classmethod
    def is_openable(cls, fnm: PathOrStr) -> bool:
        """
        Gets if a file can be opened

        Args:
            fnm (PathOrStr): file path

        Returns:
            bool: True if file can be opened; Otherwise, False
        """
        try:
            p = cls.get_absolute_path(fnm)
            if not p.exists():
                mLo.Lo.print(f"'{fnm}' does not exist")
                return False
            if not p.is_file():
                mLo.Lo.print(f"'{fnm}' does is not a file")
                return False
            if not os.access(fnm, os.R_OK):
                mLo.Lo.print(f"'{fnm}' is not readable")
                return False
            return True
        except Exception as e:
            mLo.Lo.print(f"File is not openable: {e}")
        return False

    @classmethod
    def make_directory(cls, dir: PathOrStr) -> None:
        """
        Creates path and subpaths not existing.

        Args:
            dest_dir (PathOrStr): PathLike object
        """
        p = cls.get_absolute_path(dir)
        p.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def create_temp_file(im_format: str) -> str:
        """
        Creates a temporary file

        Args:
            im_format (str): File suffix such as txt or cfg

        Raises:
            Exception: If creation of temp file fails.

        Returns:
            str: Path to temp file.
        """
        try:
            tmp = tempfile.NamedTemporaryFile(prefix="loTemp", suffix=f".{im_format}", delete=True)
            return tmp.name
        except Exception as e:
            raise Exception("Could not create temp file") from e

    @classmethod
    def delete_file(cls, fnm: PathOrStr) -> bool:
        """
        Deletes a file

        Args:
            fnm (PathOrStr): file to delete

        Returns:
            bool: True if delete is successful; Otherwise, False
        """
        p = cls.get_absolute_path(fnm)
        os.remove(p)
        if p.exists():
            mLo.Lo.print(f"'{p}' could not be deleted")
            return False
        else:
            mLo.Lo.print(f"'{p}' deleted")
        return True

    @classmethod
    def delete_files(cls, *fnms:PathOrStr) -> bool:
        """
        Deletes files
        
        Args:
            fnms (str): one or more files to delete

        Returns:
            bool: Returns True if all file are deleted; Otherwise, False
        """
        if len(fnms) == 0:
            return False
        mLo.Lo.print()
        result = True
        for s in fnms:
            result = result and cls.delete_file(s)
        return result

    @classmethod
    def save_string(cls, fnm: PathOrStr, data: str) -> None:
        if data is None:
            raise ValueError(f"No data to save in '{fnm}'")
        try:
            p = cls.get_absolute_path(fnm)
            with open(p, "w") as file:
                file.write(data)
            mLo.Lo.print(f"Saved string to file: {p}")
        except Exception as e:
            raise Exception(f"Could not save string to file: {p}") from e

    @classmethod
    def save_bytes(cls, fnm: PathOrStr, b: bytes) -> None:
        if b is None:
            raise ValueError(f"'b' is null. No data to save in '{fnm}'")
        try:
            p = cls.get_absolute_path(fnm)
            with open(p, "b") as file:
                file.write(b)
            mLo.Lo.print(f"Saved bytes to file: {p}")
        except Exception as e:
            raise Exception(f"Could not save bytes to file: {fnm}") from e

    @classmethod
    def save_array(cls, fnm: PathOrStr, arr: List[list]) -> None:
        """
        Saves a 2d array to a file as tab delimited data.

        Args:
            fnm (PathOrStr): file to save data to
            arr (List[list]): 2d array of data.
        """
        if arr is None:
            raise ValueError("'arr' is null. No data to save in '{fnm}'")
        num_rows = len(arr)
        if num_rows == 0:
            mLo.Lo.print("No data to save in '{fnm}'")
            return
        try:
            p = cls.get_absolute_path(fnm)
            with open(p, "w") as file:
                for j in range(num_rows):
                    line = "\t".join([str(v) for v in arr[j]])
                    file.write(line)
                    file.write("\n")
            mLo.Lo.print(f"Save array to file: {p}")
        except Exception as e:
            raise Exception(f"Could not save array to file: {fnm}") from e

    @classmethod
    def append_to(cls, fnm: PathOrStr, msg: str) -> None:
        """
        Appends text to a file

        Args:
            fnm (PathOrStr): File to append text to.
            msg (str): Text to append.

        Raises:
            Exception: If unable to append text.
        """
        try:
            p = cls.get_absolute_path(fnm)
            with open(p, "a") as file:
                file.write(msg)
                file.write("\n")
        except Exception as e:
            raise Exception(f"unable to append to '{fnm}'") from e

    # endregion ---------- file creation / deletion --------------------

    # region ------------- zip access ----------------------------------
    @classmethod
    def zip_access(cls, fnm: PathOrStr) -> XZipFileAccess:
        return mLo.Lo.create_instance_mcf(
            XZipFileAccess, "com.sun.star.packages.zip.ZipFileAccess", (cls.fnm_to_url(fnm),)
        )

    @classmethod
    def zip_list_uno(cls, fnm: PathOrStr) -> None:
        """Use zip_list method"""
        # replaced by more detailed Java version; see below
        zfa: XNameAccess = cls.zip_access(fnm)
        names = zfa.getElementNames()
        mLo.Lo.print(f"\nZippendContents of '{fnm}'")
        mLo.Lo.print_names(names, 1)

    @staticmethod
    def unzip_file(zfa: XZipFileAccess, fnm: PathOrStr) -> None:
        """
        Unzip File. Not yet imeplemented

        Args:
            zfa (XZipFileAccess): Zip File Access
            fnm (PathOrStr): File path

        Raises:
            NotImplementedError:
        """
        # TODO: implement unzip_file
        raise NotImplementedError

    @staticmethod
    def read_lines(in_stream: XInputStream) -> List[str] | None:
        """
        Converts a input stream to a list of strings.

        Args:
            in_stream (XInputStream): Input stream

        Returns:
            List[str] | None: If text was found in input stream the list of string; Otherwise, None
        """
        lines = []
        try:
            tis = mLo.Lo.create_instance_mcf(XTextInputStream, "com.sun.star.io.TextInputStream")
            sink = mLo.Lo.qi(XActiveDataSink, tis)
            sink.setInputStream(in_stream)

            while tis.isEOF() is False:
                lines.append(tis.readLine())
            tis.closeInput()
        except Exception as e:
            mLo.Lo.print(e)
        if len(lines) == 0:
            return None
        return lines

    @classmethod
    def get_mime_type(cls, zfa: XZipFileAccess) -> str | None:
        """
        Gets mime type for zip file access

        Args:
            zfa (XZipFileAccess): zip file access

        Raises:
            Exception: If error getting mimetype

        Returns:
            str | None: Mimetype if found; Otherwise, None
        """
        try:
            in_stream = zfa.getStreamByPattern("mimetype")
            lines = cls.read_lines(in_stream)
            if lines is not None:
                return lines[0].strip()
        except UnoException as e:
            raise Exception("Unable to get mime type") from e
        mLo.Lo.print("No mimetype found")
        return None

    # endregion ------------- zip access -------------------------------

    # region ------------- switch to Python's zip APIs -----------------

    @classmethod
    def zip_list(cls, fnm: PathOrStr) -> None:
        """
        Prints info to console for a give zip file.

        Args:
            fnm (PathOrStr): Path to zip file.
        """
        try:
            p = cls.get_absolute_path(fnm)
            with zipfile.ZipFile(p, "r") as zip:
                for info in zip.getinfo():
                    mLo.Lo.print(info.filename)
                    mLo.Lo.print("\tModified:\t" + str(datetime.datetime(*info.date_time)))
                    mLo.Lo.print("\tSystem:\t\t" + str(info.create_system) + "(0 = Windows, 3 = Unix)")
                    mLo.Lo.print("\tZIP version:\t" + str(info.create_version))
                    mLo.Lo.print("\tCompressed:\t" + str(info.compress_size) + " bytes")
                    mLo.Lo.print("\tUncompressed:\t" + str(info.file_size) + " bytes")
            mLo.Lo.print()
        except Exception as e:
            mLo.Lo.print(e)
    # endregion ---------- switch to Python's zip APIs -----------------