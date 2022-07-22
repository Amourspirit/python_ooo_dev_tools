# coding: utf-8
from __future__ import annotations
import io
import base64
from typing import TYPE_CHECKING
from PIL import Image  # LibreOffice has PHL Module.
import uno

if TYPE_CHECKING:
    from com.sun.star.graphic import XGraphic

from ..utils import file_io as mFileIO
from . import images_lo as mImgLo

from ..utils.type_var import PathOrStr


class Images(mImgLo.ImagesLo):
    @staticmethod
    def load_image(fnm: PathOrStr) -> Image.Image:
        img = None
        try:
            img = Image.open(fnm)
            img.load()
            print(f"Loaded image: '{fnm}'")
        except Exception as e:
            print(f"Unable to load '{fnm}':")
            print(f"    {e}")
        return img

    @staticmethod
    def save_image(im: Image.Image, fnm: PathOrStr) -> None:
        """
        Saves Image

        Args:
            im (Image): Image data
            fnm (PathOrStr): save path

        Raises:
            TypeError: If im is None
            Exception: If unable to save image
        """
        if im is None:
            raise TypeError("im is None")
        try:
            im.save(fp=fnm, format="png")
            print(f"Saved image to file: {fnm}")
        except Exception as e:
            raise Exception(f"Could not save image to '{fnm}'") from e

    @staticmethod
    def im_to_bytes(im: Image.Image) -> bytes:
        """
        Gets Image Bytes

        Args:
            im (Image): Image

        Returns:
            bytes: Image bytes
        """
        buf = io.BytesIO()
        im.save(buf, format="png")
        byte_im = buf.getvalue()
        return byte_im

    @classmethod
    def im_to_string(cls, im: Image.Image) -> str:
        """
        Converts image to string

        Args:
            im (Image): image

        Raises:
            Exception: If unable to convet to string

        Returns:
            str: Image as string
        """
        try:
            b = cls.im_to_bytes(im)
            s = base64.b64encode(b).decode("utf-8")
            return s
        except Exception as e:
            raise Exception("Converting image to string is not possible:") from e

    @staticmethod
    def string_to_im(s: str) -> Image.Image:
        try:
            imgdata = base64.b64decode(s)
            image = Image.open(io.BytesIO(imgdata))
            return image
        except Exception as e:
            print("Converting string to image is not possible:")
            print(f"    {e}")
            return None

    @staticmethod
    def bytes_to_im(b: bytes) -> Image.Image:
        try:
            image = Image.open(io.BytesIO(b))
            return image
        except Exception as e:
            print("Converting bytes to image is not possible:")
            print(f"    {e}")
            return None

    @classmethod
    def im_to_graphic(cls, im: Image.Image) -> XGraphic | None:
        if im is None:
            print("No image found")
            return None

        temp_fnm = mFileIO.FileIO.create_temp_file(im_format="png")
        if temp_fnm is None:
            print("Could not create a temporary file for the image")
            return None

        im.save(temp_fnm, format="png")
        graphic = cls.load_graphic_file(temp_fnm)
        mFileIO.FileIO.delete_file(temp_fnm)
        return graphic

    @classmethod
    def graphic_to_im(cls, graphic: XGraphic) -> Image.Image | None:
        if graphic is None:
            print("No graphic found")
            return None
        tmp_fnm = mFileIO.FileIO.create_temp_file("png")
        if tmp_fnm is None:
            return None
        cls.save_graphic(graphic, tmp_fnm, "png")
        im = cls.load_image(tmp_fnm)
        mFileIO.FileIO.delete_file(tmp_fnm)
        return im
