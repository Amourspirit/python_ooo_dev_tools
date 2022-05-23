# coding: utf-8
from __future__ import annotations
from PIL import Image  # LibreOffice has PHL Module.
import io
import base64
from typing import TYPE_CHECKING, Tuple
import sys
import uno
from com.sun.star.awt import Size
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNameContainer
from com.sun.star.document import XMimeTypeInfo
from com.sun.star.graphic import XGraphicProvider

if TYPE_CHECKING:
    from com.sun.star.graphic import XGraphic

from ..utils import lo as mLo
from ..utils import file_io as mFileIO
from ..utils import props as mProps

# if sys.version_info >= (3, 10):
#     from typing import Union
# else:
#     from typing_extensions import Union


# Lo = m_lo.Lo
# FileIO = m_file_io.FileIO
# Props = m_props.Props


class Images:
    @staticmethod
    def get_bitmap(fnm: str) -> str | None:
        try:
            bitmap_container = mLo.Lo.create_instance_msf(XNameContainer, "com.sun.star.drawing.BitmapTable")
            if not mFileIO.FileIO.is_openable(fnm):
                return None

            pic_url = mFileIO.FileIO.fnm_to_url(fnm)
            if pic_url is None:
                return None
            bitmap_container.insertByName(fnm, pic_url)
            # use the filename as the name of the bitmap

            # return the bitmap as a string
            str(bitmap_container.getByName(fnm))
        except Exception as e:
            print(f"Could not create a bitmap container for '{fnm}'")
            print(f"    {e}")
            return None

    @staticmethod
    def load_image(fnm: str) -> Image.Image | None:
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
    def save_image(im: Image.Image, fnm: str) -> None:
        if im is None:
            print(f"No data to save in '{fnm}'")
            return
        try:
            im.save(fp=fnm, format="png")
            print(f"Saved image to file: {fnm}")
        except Exception as e:
            print(f"Could not save image to '{fnm}'")
            print(f"    {e}")

    @staticmethod
    def im_to_bytes(im: Image.Image) -> bytes:
        buf = io.BytesIO()
        im.save(buf, format="png")
        byte_im = buf.getvalue()
        return byte_im

    @classmethod
    def im_to_string(cls, im: Image.Image) -> str | None:
        try:
            b = cls.im_to_bytes(im)
            s = base64.b64encode(b).decode("utf-8")
            return s
        except Exception as e:
            print("Converting image to string is not possilbe:")
            print(f"    {e}")
            return None

    @staticmethod
    def string_to_im(s: str) -> Image.Image:
        try:
            imgdata = base64.b64decode(s)
            image = Image.open(io.BytesIO(imgdata))
            return image
        except Exception as e:
            print("Converting string to image is not possilbe:")
            print(f"    {e}")
            return None

    @staticmethod
    def bytes_to_im(b: bytes) -> Image.Image:
        try:
            image = Image.open(io.BytesIO(b))
            return image
        except Exception as e:
            print("Converting bytes to image is not possilbe:")
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

    @staticmethod
    def load_graphic_file(im_fnm: str) -> XGraphic | None:
        print(f"Loading XGraphic from '{im_fnm}'")
        gprovider = mLo.Lo.create_instance_mcf(XGraphicProvider, "com.sun.star.graphic.GraphicProvider")
        if gprovider is None:
            print("Graphic Provider could not be found")
            return None

        file_props = mProps.Props.make_props(URL=mFileIO.FileIO.fnm_to_url(im_fnm))
        try:
            return gprovider.queryGraphic(file_props)
        except Exception as e:
            print(f"Could not load XGraphic from '{im_fnm}'")
            print(f"    {e}")
        return None

    @classmethod
    def get_size_pixels(cls, im_fnm) -> Size | None:
        graphic = cls.load_graphic_file(im_fnm)
        if graphic is None:
            return None
        return mProps.Props.get_property(xprops=graphic, name="SizePixel")

    @classmethod
    def get_size_100mm(cls, im_fnm: str) -> Size | None:
        graphic = cls.load_graphic_file(im_fnm)
        if graphic is None:
            return None
        return mProps.Props.get_property(xprops=graphic, name="Size100thMM")

    @staticmethod
    def load_graphic_link(graphic_link: object) -> XGraphic | None:
        gprovider = mLo.Lo.create_instance_mcf(XGraphicProvider, "com.sun.star.graphic.GraphicProvider")
        if gprovider is None:
            print("Graphic Provider could not be found")
            return None

        xprops = mLo.Lo.qi(XPropertySet, graphic_link)
        if xprops is None:
            return None
        gprops = mProps.Props.make_props(URL=str(xprops.getPropertyValue("GraphicURL")))

        try:
            return gprovider.queryGraphic(gprops)
        except Exception as e:
            print(f"Unable to retrieve graphic")
            print(f"    {e}")
        return None

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

    @staticmethod
    def save_graphic(pic: XGraphic, fnm: str, im_format: str) -> None:
        print(f"Saving graphic in '{fnm}'")

        if pic is None:
            print("Supplied image is null")
            return

        gprovider = mLo.Lo.create_instance_mcf(XGraphicProvider, "com.sun.star.graphic.GraphicProvider")
        if gprovider is None:
            print("Graphic Provider could not be found")
            return

        png_props = mProps.Props.make_props(URL=mFileIO.FileIO.fnm_to_url(fnm), MimeType=f"image/{im_format}")

        try:
            gprovider.storeGraphic(png_props)
        except Exception as e:
            print("Unable to save graphic")
            print(f"    {e}")

    @staticmethod
    def get_mime_types() -> Tuple[str, ...]:
        mi = mLo.Lo.create_instance_mcf(XMimeTypeInfo, "com.sun.star.drawing.GraphicExportFilter")
        return mi.getSupportedMimeTypeNames()

    @classmethod
    def change_to_mime(cls, im_format: str) -> str:
        names = cls.get_mime_types()
        imf = im_format.lower().strip()
        for name in names:
            if imf in name:
                print(f"using mime type: {name}")
                return name
        print("No matching mime type, so using image/png")
        return "image/png"

    @classmethod
    def calc_scale(cls, fnm: str, max_width: int, max_height: int) -> Size | None:
        """
        Calculate a new size for the image in fnm that is no bigger than
        maxWidth x maxHeight mm's
        This involves a rescaling of the image so it is not distorted.
        The new size is returned in mm units
        """
        im_size = cls.get_size_100mm(fnm)  # in 1/100 mm units
        if im_size is None:
            return None
        # calculate the scale factors to obtain these maximums
        width_scale = (max_width * 100) / im_size.Width
        height_scale = (max_height * 100) / im_size.Height

        # use the smallest scale factor
        scale_factor = min(width_scale, height_scale)

        # calculate new dimensions for the image
        w = round(im_size.Width * scale_factor / 100)
        h = round(im_size.Height * scale_factor / 100)
        return Size(w, h)
