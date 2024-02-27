# coding: utf-8
from __future__ import annotations
from dataclasses import dataclass
from typing import Any, TYPE_CHECKING, ByteString, Tuple, cast, overload
import base64
import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNameContainer
from com.sun.star.document import XMimeTypeInfo
from com.sun.star.graphic import XGraphicProvider
from ooo.dyn.awt.size import Size as UnoSize

from ooodev.utils.data_type.width_height_fraction import WidthHeightFraction
from ooodev.exceptions import ex as mEx
from ooodev.units.unit_convert import UnitConvert, UnitLength
from ooodev.utils import file_io as mFileIO
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.data_type.size import Size

if TYPE_CHECKING:
    from com.sun.star.graphic import XGraphic
    from com.sun.star.awt import XBitmap
    from ooodev.utils.type_var import PathOrStr
else:
    PathOrStr = Any

# see Also: https://ask.libreoffice.org/t/graphicurl-no-longer-works-in-6-1-0-3/35459/3
# see Also: https://tomazvajngerl.blogspot.com/2018/03/improving-image-handling-in-libreoffice.html


@dataclass
class BitmapArgs:
    """Args use when calling bitmap method that support args"""

    name: str
    """Name given to bitmap in the ``BitmapTable`` container"""
    auto_name: bool = True
    """If True then name is checked. If it exist in ``BitmapTable`` then a new unique name is created"""
    auto_update: bool = False
    """If True and name already exist in ``BitmapTable`` container then the existing name is updated. ``auto_name`` must be ``False``"""
    out_name: str = ""
    """The name used in the ``BitmapTable`` container will be assigned to the property."""


class ImagesLo:
    @classmethod
    def get_bitmap(cls, fnm: PathOrStr, args: BitmapArgs | None = None) -> XBitmap:
        """
        Gets an image from path

        Args:
            fnm (PathOrStr): path to image
            args (BitmapArgs, optional): Args for the method

        Raises:
            UnOpenableError: If image is not able to be opened
            Exception: If unable to get bitmap

        Returns:
            XBitmap: Bitmap

        Note:
            ``args`` are in and out values. When ``args`` are used the ``args.out_name`` is assigned the name used to add bitmap to ``BitmapTable``.

            If ``args.auto_name`` is ``False`` then ``args.out_name`` will be the same name assigned to ``args.name``;
            Otherwise, if ``args.auto_name`` is ``True`` then ``args.out_name`` will be then be value of ``args.name`` with a number appended.
            For example a name of ``Bitmap `` would typically return a ``out_name`` such as ``Bitmap 1`` or ``Bitmap 2``, etc.


        .. versionchanged:: 0.9.0
            Added ``args`` to method.
        """
        # sourcery skip: raise-specific-error
        try:
            fp = mFileIO.FileIO.get_absolute_path(fnm)
            if not mFileIO.FileIO.is_openable(fp):
                raise mEx.UnOpenableError(fnm=fp)
            btc = mLo.Lo.create_instance_msf(XNameContainer, "com.sun.star.drawing.BitmapTable", raise_err=True)
            if args is not None:
                if args.auto_name:
                    name = cls._get_unique_container_el_name(args.name, btc)
                else:
                    name = args.name
                args.out_name = name
            else:
                # Get a sequence name Bitmap 1, Bitmap 2 etc
                name = cls._get_unique_container_el_name("Bitmap ", btc)
            pic_url = mFileIO.FileIO.fnm_to_url(fp)

            if args is not None and args.auto_update and btc.hasByName(name):
                btc.replaceByName(name, pic_url)
            elif (
                args is not None
                and args.auto_update
                and not btc.hasByName(name)
                or (args is None or not args.auto_update)
                and not btc.hasByName(name)
            ):
                btc.insertByName(name, pic_url)
            else:
                return btc.getByName(name)
            # return bitmap.getDIB() # byte sequence
            return btc.getByName(name)
        except mEx.UnOpenableError:
            raise
        except Exception as e:
            raise Exception(f"Could not create a bitmap container for '{fnm}'") from e

    @classmethod
    def get_bitmap_from_b64(
        cls, value: ByteString, args: BitmapArgs | None = None, img_format: str = "png"
    ) -> XBitmap:
        """
        Gets a ``XBitmap`` from a base 64 encoded byte string.

        Args:
            value (ByteString): Base 64 image
            args (BitmapArgs, optional): Args for the method
            img_format (str, optional): Image format such as ``png``, ``jpeg``

        Returns:
            XBitmap: Image as ``XBitmap``

        .. versionadded:: 0.9.0
        """
        tmp = mFileIO.FileIO.create_temp_file(img_format)
        try:
            with open(tmp, "wb") as fh:
                fh.write(base64.decodebytes(value))  # type: ignore
            return cls.get_bitmap(tmp, args)
        finally:
            mFileIO.FileIO.delete_file(tmp)

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
            provider = mLo.Lo.create_instance_mcf(
                XGraphicProvider, "com.sun.star.graphic.GraphicProvider", raise_err=True
            )
            file_props = mProps.Props.make_props(URL=mFileIO.FileIO.fnm_to_url(fnm))
            result = provider.queryGraphic(file_props)  # type: ignore
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
            ~ooodev.utils.data_type.size.Size: Size containing Width and Height
        """
        graphic = cls.load_graphic_file(im_fnm)
        sz = cast(UnoSize, mProps.Props.get(graphic, "SizePixel"))
        return Size(sz.Width, sz.Height)

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
            ~ooodev.utils.data_type.size.Size: Size containing Width and Height
        """
        graphic = cls.load_graphic_file(im_fnm)
        sz = cast(UnoSize, mProps.Props.get(graphic, "Size100thMM"))
        return Size(sz.Width, sz.Height)

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
        # sourcery skip: raise-specific-error
        xprops = mLo.Lo.qi(XPropertySet, graphic_link, True)

        try:
            graphic = xprops.getPropertyValue("Graphic")
            if graphic is None:
                raise Exception("Graphic is None")
            return graphic
        except Exception as e:
            raise Exception("Unable to retrieve graphic") from e

    # region save_graphic()
    @overload
    @staticmethod
    def save_graphic(pic: XGraphic, fnm: PathOrStr) -> None: ...

    @overload
    @staticmethod
    def save_graphic(pic: XGraphic, fnm: PathOrStr, im_format: str) -> None: ...

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
                ext = mInfo.Info.get_ext(fnm)
                im_format = "" if ext is None else ext
                if not im_format:
                    raise ValueError(
                        "Unable to get image format from fnm. Does fnm have an file extension such as myfile.png?"
                    )
                im_format = im_format.lower()

            provider = mLo.Lo.create_instance_mcf(
                XGraphicProvider, "com.sun.star.graphic.GraphicProvider", raise_err=True
            )

            png_props = mProps.Props.make_props(URL=mFileIO.FileIO.fnm_to_url(fnm), MimeType=f"image/{im_format}")

            provider.storeGraphic(pic, png_props)  # type: ignore
        except Exception as e:
            raise mEx.ImageError(f'Error saving graphic for "{fnm}') from e

    # endregion save_graphic()

    @staticmethod
    def get_mime_types() -> Tuple[str, ...]:
        """
        Gets mime types from ``com.sun.star.drawing.GraphicExportFilter``

        Returns:
            Tuple[str, ...]: Tuple of mime types such as ``('image/png', 'image/jpeg', 'application/pdf', ...)``
        """
        mi = mLo.Lo.create_instance_mcf(XMimeTypeInfo, "com.sun.star.drawing.GraphicExportFilter", raise_err=True)
        return mi.getSupportedMimeTypeNames()

    @classmethod
    def change_to_mime(cls, im_format: str) -> str:
        """
        Change to Mime type. If the input is valid then it is returned otherwise ``image/png`` is returned.

        Args:
            im_format (str): An expected mime type such as ``image/jpeg`` if not found then ``image/png`` is returned.

        Returns:
            str: Mime type. Defaults to ``image/png``

        See Also:
            :py:meth:`~.ImagesLo.get_mime_types`
        """
        names = cls.get_mime_types()
        imf = im_format.lower().strip()
        for name in names:
            if imf in name:
                print(f"using mime type: {name}")
                return name
        print("No matching mime type, so using image/png")
        return "image/png"

    # region internal methods
    @staticmethod
    def _get_unique_container_el_name(prefix: str, nc: XNameContainer) -> str:
        """
        Gets the next name that does not exist in the container.

        Lets say ``prefix`` is ``Transparency `` then names are search in sequence.
        ``Transparency 1``, ``Transparency 3``, ``Transparency 3``, etc until a unique name is found.

        Args:
            prefix (str): Any string such as ``Transparency ``
            nc (XNameContainer | None, optional): Container. Defaults to None.

        Returns:
            str: Unique name
        """
        names = nc.getElementNames()
        i = 1
        name = f"{prefix}{i}"
        while name in names:
            i += 1
            name = f"{prefix}{i}"
        return name

    # endregion internal methods

    # region Image Calculations
    @classmethod
    def calc_scale(cls, fnm: PathOrStr, max_width: int, max_height: int) -> Size | None:
        """
        Calculate a new size for the image in fnm that is no bigger than
        maxWidth x maxHeight mm's.
        This involves a re-scaling of the image so it is not distorted.
        The new size is returned in ``mm`` units.

        Args:
            fnm (PathOrStr): Path to image
            max_width (int): Max height.
            max_height (int): max_width

        Returns:
            ~ooodev.utils.data_type.size.Size | None:
        """
        im_size = cls.get_size_100mm(fnm)  # in 1/100 mm units
        if im_size is None:
            return None
        # calculate the scale factors to obtain these maximums
        width_scale = (max_width * 100) / im_size.width
        height_scale = (max_height * 100) / im_size.height

        # use the smallest scale factor
        scale_factor = min(width_scale, height_scale)

        # calculate new dimensions for the image
        w = round(im_size.width * scale_factor / 100)
        h = round(im_size.height * scale_factor / 100)
        return Size(w, h)

    @staticmethod
    def calc_aspect_ratio_fit(
        src_width: float, src_height: float, max_width: float, max_height: float
    ) -> WidthHeightFraction:
        """
        Conserve aspect ratio of the original region. Useful when shrinking/enlarging
        images to fit into a certain area.

        Args:
            src_width (float): Source Width.
            src_height (float): Source height.
            max_width (float): Max width.
            max_height (float): Max Height.

        Returns:
            WidthHeightFraction:
        """
        ratio = min(max_width / src_width, max_height / src_height)
        return WidthHeightFraction(width=src_width * ratio, height=src_height * ratio)

    @staticmethod
    def calc_scale_factor(length: float, new_length: float) -> float:
        """
        Gets the scale factor between ``length`` and ``new_length``

        Args:
            length (float): Length, usually width or height of an image.
            new_length (float): The new length.

        Returns:
            float: Scale Factor.

        Example:

        .. code-block:: python

            >>> print(ImagesLo.calculate_scale_factor(50.0, 150.0))
            300.0

        Note:
            If ``length`` or ``new_length`` is ``0`` then ``1.0`` is returned.
        """
        return 1.0 if length == 0 or new_length == 0 else (new_length / length) * 100

    @classmethod
    def calc_scale_crop(cls, orig_len: float, new_len: float, start_crop: float, end_crop: float) -> float:
        """
        Gets the scale factor for cropping values with keep size.

        Args:
            orig_len (float): Original length. Usually image original height or width.
            new_len (float): The new length. Usually the new image height or width.
            start_crop (float): Start Crop. Usually the amount to crop image left or top.
            end_crop (float): Start Crop. Usually the amount to crop image right or bottom.

        Returns:
            float: Scale Factor.
        """
        new_length = start_crop + end_crop + new_len
        return cls.calc_scale_factor(orig_len, new_length)

    @staticmethod
    def calc_keep_scale_len(orig_len: float, start_crop: float, end_crop: float, scale: float) -> float:
        """
        Gets the length (usually width or height) for cropping values with keep scale.

        Args:
            orig_len (float): Original length. Usually image original height or width.
            start_crop (float): Start Crop. Usually the amount to crop image left or top.
            end_crop (float): Start Crop. Usually the amount to crop image right or bottom.
            scale (float): Scale Factor.

        Returns:
            float: Length that usually represents the new image width or height.
        """
        img_len = orig_len - (start_crop + end_crop)
        return img_len * scale

    # endregion Image Calculations

    @staticmethod
    def get_dpi_width_height(width: int, height: int, resolution: int) -> tuple[int, int]:
        """
        Gets the DPI of the image.

        Args:
            width (int): Width in ``1/100 mm``.
            height (int): Height in ``1/100 mm``.
            resolution (int): Resolution in dpi.

        Returns:
            tuple[int, int]: Width and Height that represents number of pixels to create the resolution.

        .. versionadded:: 0.21.3
        """
        width_in = UnitConvert.convert(width, frm=UnitLength.MM100, to=UnitLength.IN)
        height_in = UnitConvert.convert(height, frm=UnitLength.MM100, to=UnitLength.IN)

        dpi_width = round(resolution * width_in)
        dpi_height = round(resolution * height_in)
        return dpi_width, dpi_height
