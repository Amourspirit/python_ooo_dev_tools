# coding: utf-8
from __future__ import annotations
from typing import TYPE_CHECKING, Tuple, overload
import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNameContainer
from com.sun.star.document import XMimeTypeInfo
from com.sun.star.graphic import XGraphicProvider

if TYPE_CHECKING:
    from com.sun.star.graphic import XGraphic
    from com.sun.star.awt import XBitmap

from ooo.dyn.awt.size import Size

from ..utils import lo as mLo
from ..utils import file_io as mFileIO
from ..utils import props as mProps
from ..utils import info as mInfo
from ..exceptions import ex as mEx
from ..utils.type_var import PathOrStr

# see Also: https://ask.libreoffice.org/t/graphicurl-no-longer-works-in-6-1-0-3/35459/3
# see Also: https://tomazvajngerl.blogspot.com/2018/03/improving-image-handling-in-libreoffice.html


class ImagesLo:
    @staticmethod
    def get_bitmap(fnm: PathOrStr) -> XBitmap:
        """
        Gets an image from path

        Args:
            fnm (PathOrStr): path to image

        Raises:
            UnOpenableError: If image is not able to be opened
            Exception: If unable to get bitmap

        Returns:
            XBitmap: Bitmap
        """
        try:
            fp = mFileIO.FileIO.get_absolute_path(fnm)
            bitmap_container = mLo.Lo.create_instance_msf(XNameContainer, "com.sun.star.drawing.BitmapTable")
            if not mFileIO.FileIO.is_openable(fp):
                raise mEx.UnOpenableError(fnm=fp)

            name = str(fp)
            pic_url = mFileIO.FileIO.fnm_to_url(fp)

            # use the filename as the name of the bitmap
            bitmap_container.insertByName(name, pic_url)

            bitmap = bitmap_container.getByName(name)
            # return bitmap.getDIB() # byte sequence
            return bitmap
        except mEx.UnOpenableError:
            raise
        except Exception as e:
            raise Exception(f"Could not create a bitmap container for '{fnm}'") from e

    @staticmethod
    def load_graphic_file(im_fnm: PathOrStr) -> XGraphic:
        """
        Loads a graphic file

        Args:
            im_fnm (PathOrStr): Graphic file path

        Raises:
            FileNotFoundError: If ``im_fnm`` does not exist.
            ImageError: If unable to load graphic

        Returns:
            XGraphic: Graphic
        """
        mLo.Lo.print(f"Loading XGraphic from '{im_fnm}'")
        try:
            _ = mFileIO.FileIO.is_exist_file(fnm=im_fnm, raise_err=True)
            fnm = mFileIO.FileIO.get_absolute_path(im_fnm)
            gprovider = mLo.Lo.create_instance_mcf(
                XGraphicProvider, "com.sun.star.graphic.GraphicProvider", raise_err=True
            )
            file_props = mProps.Props.make_props(URL=mFileIO.FileIO.fnm_to_url(fnm))
            result = gprovider.queryGraphic(file_props)
            if result is None:
                raise mEx.UnKnownError("None Value: queryGraphic() returned None")
            return result
        except FileNotFoundError:
            raise
        except Exception as e:
            raise mEx.ImageError(f"Could not load XGraphic from '{im_fnm}'") from e

    @classmethod
    def get_size_pixels(cls, im_fnm: PathOrStr) -> Size:
        """
        Gets an image/graphic size from image file path

        Args:
            im_fnm (PathOrStr): file path

        Returns:
            Size: Size containing Width and Height
        """
        graphic = cls.load_graphic_file(im_fnm)
        return mProps.Props.get(graphic, "SizePixel")

    @classmethod
    def get_size_100mm(cls, im_fnm: PathOrStr) -> Size:
        """
        The Size of the graphic in ``100th mm``.

        This property may not be available in case of pixel graphics or
        if the logical size can not be determined correctly for some formats
        without loading the whole graphic


        Args:
            im_fnm (PathOrStr): Path to image/graphic

        Raises:
            PropertyError: If Size100thMM property is not available

        Returns:
            Size: Size containing Width and Height
        """
        graphic = cls.load_graphic_file(im_fnm)
        return mProps.Props.get(graphic, "Size100thMM")

    @staticmethod
    def load_graphic_link(graphic_link: object) -> XGraphic:
        """
        Loads a graphic link

        Args:
            graphic_link (object): Object that implements XPropertySet interface

        Raises:
            CreateInstanceMcfError: If unable to create graphic.GraphicProvider instance
            MissingInterfaceError: If graphic_link does not implement XPropertySet interface
            Exception: If unable to load graphic

        Returns:
            XGraphic: Graphic
        """
        xprops = mLo.Lo.qi(XPropertySet, graphic_link, True)

        try:
            graphic = xprops.getPropertyValue("Graphic")
            if graphic is None:
                raise Exception("Grapich is None")
            return graphic
        except Exception as e:
            raise Exception(f"Unable to retrieve graphic") from e

    # region save_graphic()
    @overload
    @staticmethod
    def save_graphic(pic: XGraphic, fnm: PathOrStr) -> None:
        ...

    @overload
    @staticmethod
    def save_graphic(pic: XGraphic, fnm: PathOrStr, im_format: str) -> None:
        ...

    @staticmethod
    def save_graphic(pic: XGraphic, fnm: PathOrStr, im_format: str = "") -> None:
        """
        Save a graphic to disk

        Args:
            pic (XGraphic): Graphic object
            fnm (PathOrStr): File path to save graphic to
            im_format (str, optional): Image format such as ``png``. Defaults to extension of ``fnm``.

        Raises:
            ImageError: If error occurs.
        """
        mLo.Lo.print(f"Saving graphic in '{fnm}'")

        try:
            if pic is None:
                raise TypeError("Expected pic to be XGraphic instance but got None")
            if not im_format:
                im_format = mInfo.Info.get_ext(fnm)
                if not im_format:
                    raise ValueError(
                        "Unable to get image format from fnm. Does fnm have an file extension such as myfile.png?"
                    )
                im_format = im_format.lower()

            gprovider = mLo.Lo.create_instance_mcf(
                XGraphicProvider, "com.sun.star.graphic.GraphicProvider", raise_err=True
            )

            png_props = mProps.Props.make_props(URL=mFileIO.FileIO.fnm_to_url(fnm), MimeType=f"image/{im_format}")

            gprovider.storeGraphic(pic, png_props)
        except Exception as e:
            raise mEx.ImageError(f'Error saving graphic for "{fnm}') from e

    # end region save_graphic()

    @staticmethod
    def get_mime_types() -> Tuple[str, ...]:
        mi = mLo.Lo.create_instance_mcf(XMimeTypeInfo, "com.sun.star.drawing.GraphicExportFilter", raise_err=True)
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
    def calc_scale(cls, fnm: PathOrStr, max_width: int, max_height: int) -> Size | None:
        """
        Calculate a new size for the image in fnm that is no bigger than
        maxWidth x maxHeight mm's
        This involves a re-scaling of the image so it is not distorted.
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
