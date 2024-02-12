from __future__ import annotations
from typing import TYPE_CHECKING
import contextlib
import uno


if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties
    from com.sun.star.util import Color  # type def
    from com.sun.star.awt import XBitmap
    from com.sun.star.drawing.BitmapMode import BitmapModeProto  # type: ignore
    from com.sun.star.drawing.RectanglePoint import RectanglePointProto  # type: ignore
    from com.sun.star.drawing.FillStyle import FillStyleProto  # type: ignore
    from com.sun.star.awt import Gradient  # Struct
    from com.sun.star.drawing import Hatch  # Struct
    from com.sun.star.text import GraphicCrop  # Struct


class FillPropertiesPartial:
    """
    Partial class for XStyleFamiliesSupplier.

    See Also:
        `API FillProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1FillProperties.html>`_
    """

    def __init__(self, component: FillProperties) -> None:
        """
        Constructor

        Args:
            component (FillProperties): UNO Component that implements ``com.sun.star.drawing.FillProperties`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``FillProperties``.
        """
        self.__component = component

    # region FillProperties
    @property
    def fill_background(self) -> bool:
        """
        Gets/Sets whether the transparent background of a hatch filled area is drawn in the current background color.

        If this is ``True``, the transparent background of a hatch filled area is drawn in the current background color.
        """
        return self.__component.FillBackground

    @fill_background.setter
    def fill_background(self, value: bool) -> None:
        self.__component.FillBackground = value

    @property
    def fill_bitmap(self) -> XBitmap | None:
        """
        Gets/Sets the bitmap used for filling.

        If the property ``fill_style`` is set to ``FillStyle.BITMAP``, this is the bitmap used.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.FillBitmap
        return None

    @fill_bitmap.setter
    def fill_bitmap(self, value: XBitmap) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.FillBitmap = value

    @property
    def fill_bitmap_logical_size(self) -> bool:
        """
        Gets/Sets if the size is given in percentage or as an absolute value.

        If this is ``True``, the properties ``fill_bitmap_size_x`` and ``fill_bitmap_size_y`` contain the size of the tile in percent of the size of the original bitmap.
        If this is ``False``, the size of the tile is specified with ``1/100th mm``.
        """
        return self.__component.FillBitmapLogicalSize

    @fill_bitmap_logical_size.setter
    def fill_bitmap_logical_size(self, value: bool) -> None:
        self.__component.FillBitmapLogicalSize = value

    @property
    def fill_bitmap_mode(self) -> BitmapModeProto:
        """
        Gets/Sets how an area is filled with a single bitmap.

        This enum selects how an area is filled with a single bitmap.

        This property corresponds to the properties ``fill_bitmap_stretch`` and ``fill_bitmap_tile``.

        If set to ``BitmapMode.REPEAT``, the property ``fill_bitmap_stretch`` is set to ``False``, and the property ``fill_bitmap_tile`` is set to ``True``.

        If set to ``BitmapMode.STRETCH``, the property ``fill_bitmap_stretch`` is set to ``True``, and the property ``fill_bitmap_tile`` is set to ``False``.

        If set to ``BitmapMode.NO_REPEAT``, both properties ``fill_bitmap_stretch`` and ``fill_bitmap_tile`` are set to ``False``.
        """
        return self.__component.FillBitmapMode

    @fill_bitmap_mode.setter
    def fill_bitmap_mode(self, value: BitmapModeProto) -> None:
        self.__component.FillBitmapMode = value

    @property
    def fill_bitmap_name(self) -> str:
        """
        If the property FillStyle is set to ``FillStyle.BITMAP``, this is the name of the used fill bitmap style.
        """
        return self.__component.FillBitmapName

    @fill_bitmap_name.setter
    def fill_bitmap_name(self, value: str) -> None:
        self.__component.FillBitmapName = value

    @property
    def fill_bitmap_offset_x(self) -> int:
        """
        Gets/Sets - Every second line of tiles is moved the given percent of the width of the bitmap.
        """
        return self.__component.FillBitmapOffsetX

    @fill_bitmap_offset_x.setter
    def fill_bitmap_offset_x(self, value: int) -> None:
        self.__component.FillBitmapOffsetX = value

    @property
    def fill_bitmap_offset_y(self) -> int:
        """
        Gets/Sets - Every second row of tiles is moved the given percent of the height of the bitmap.
        """
        return self.__component.FillBitmapOffsetY

    @fill_bitmap_offset_y.setter
    def fill_bitmap_offset_y(self, value: int) -> None:
        self.__component.FillBitmapOffsetY = value

    @property
    def fill_bitmap_position_offset_x(self) -> int:
        """
        Gets/Sets the horizontal offset where the tile starts.

        It is given in percent in relation to the width of the bitmap.
        """
        return self.__component.FillBitmapPositionOffsetX

    @fill_bitmap_position_offset_x.setter
    def fill_bitmap_position_offset_x(self, value: int) -> None:
        self.__component.FillBitmapPositionOffsetX = value

    @property
    def fill_bitmap_position_offset_y(self) -> int:
        """
        Gets/Sets the vertical offset where the tile starts.

        It is given in percent in relation to the height of the bitmap.
        """
        return self.__component.FillBitmapPositionOffsetY

    @fill_bitmap_position_offset_y.setter
    def fill_bitmap_position_offset_y(self, value: int) -> None:
        self.__component.FillBitmapPositionOffsetY = value

    @property
    def fill_bitmap_rectangle_point(self) -> RectanglePointProto:
        """
        Gets/Sets - RectanglePoint specifies the position inside of the bitmap to use as the top left position for rendering.
        """
        return self.__component.FillBitmapRectanglePoint

    @fill_bitmap_rectangle_point.setter
    def fill_bitmap_rectangle_point(self, value: RectanglePointProto) -> None:
        self.__component.FillBitmapRectanglePoint = value

    @property
    def fill_bitmap_size_x(self) -> int:
        """
        Gets/Sets the width of the tile for filling.

        Depending on the property ``fill_bitmap_logical_size``, this is either relative or absolute.

        If ``fill_bitmap_logical_size`` is ``True`` then property contain the size of the tile in percent
        of the size of the original bitmap; Otherwise, the size of the tile is specified with 1/100th mm.
        """
        # percentage or 1/100th mm
        return self.__component.FillBitmapSizeX

    @fill_bitmap_size_x.setter
    def fill_bitmap_size_x(self, value: int) -> None:
        self.__component.FillBitmapSizeX = value

    @property
    def fill_bitmap_size_y(self) -> int:
        """
        Gets/Sets the height of the tile for filling.

        This is the height of the tile for filling.

        Depending on the property FillBitmapLogicalSize, this is either relative or absolute.

        If ``fill_bitmap_logical_size`` is ``True`` then property contain the size of the tile in percent
        of the size of the original bitmap; Otherwise, the size of the tile is specified with 1/100th mm.
        """
        return self.__component.FillBitmapSizeY

    @fill_bitmap_size_y.setter
    def fill_bitmap_size_y(self, value: int) -> None:
        self.__component.FillBitmapSizeY = value

    @property
    def fill_bitmap_stretch(self) -> bool | None:
        """
        Gets/Sets if the fill bitmap is stretched to fill the area of the shape.

        This property should not be used anymore and is included here for completeness.
        The ``fill_bitmap_mode`` property can be used instead to set all supported bitmap modes.

        If set to ``True``, the value of the ``fill_bitmap_mode`` property changes to ``BitmapMode.STRETCH``.
        BUT: behavior is undefined, if the property ``fill_bitmap_tile`` is ``True`` too.

        If set to ``False``, the value of the ``fill_bitmap_mode`` property changes to ``BitmapMode.REPEAT``
        or ``BitmapMode.NO_REPEAT``, depending on the current value of the ``fill_bitmap_tile`` property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.FillBitmapStretch
        return None

    @fill_bitmap_stretch.setter
    def fill_bitmap_stretch(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.FillBitmapStretch = value

    @property
    def fill_bitmap_tile(self) -> bool | None:
        """
        Get/Sets the fill bitmap is repeated to fill the area of the shape.

        This property should not be used anymore and is included here for completeness.
        The ``fill_bitmap_mode`` property can be used instead to set all supported bitmap modes.

        If set to ``True``, the value of the ``fill_bitmap_mode`` property changes to ``BitmapMode.REPEAT``.
        BUT: behavior is undefined, if the property ``fill_bitmap_stretch`` is ``True`` too.

        If set to ``False``, the value of the ``fill_bitmap_mode`` property changes to ``BitmapMode.STRETCH``
        or ``BitmapMode.NO_REPEAT``, depending on the current value of the ``fill_bitmap_stretch`` property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.FillBitmapTile
        return None

    @fill_bitmap_tile.setter
    def fill_bitmap_tile(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.FillBitmapTile = value

    @property
    def fill_bitmap_url(self) -> str | None:
        """
        Gets/Sets the URL of the bitmap used for filling.

        If the property ``fill_style`` is set to ``FillStyle.BITMAP``, this is a URL to the bitmap used.

        Note the new behavior since it this was deprecated: This property can only be set and only external URLs are supported (no more vnd.sun.star.GraphicObject scheme).
        When a URL is set, then it will load the bitmap and set the ``fill_bitmap`` property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.FillBitmapURL
        return None

    @fill_bitmap_url.setter
    def fill_bitmap_url(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.FillBitmapURL = value

    @property
    def fill_color(self) -> Color:
        """
        Gets/Sets the color used for filling.

        If the property ``fill_style`` is set to ``FillStyle.SOLID``, this is the color used.
        """
        return self.__component.FillColor

    @fill_color.setter
    def fill_color(self, value: Color) -> None:
        self.__component.FillColor = value

    @property
    def fill_gradient(self) -> Gradient | None:
        """
        Gets/Sets the gradient used for filling.

        If the property ``fill_style`` is set to ``FillStyle.GRADIENT``, this describes the gradient used.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.FillGradient
        return None

    @fill_gradient.setter
    def fill_gradient(self, value: Gradient) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.FillGradient = value

    @property
    def fill_gradient_name(self) -> str:
        """
        Gets/Sets the name of the used fill gradient style.

        If the property ``fill_style`` is set to ``FillStyle.GRADIENT``, this is the name of the used fill gradient style.
        """
        return self.__component.FillGradientName

    @fill_gradient_name.setter
    def fill_gradient_name(self, value: str) -> None:
        self.__component.FillGradientName = value

    @property
    def fill_hatch(self) -> Hatch | None:
        """
        Gets/Sets the hatch used for filling.

        If the property ``fill_style`` is set to ``FillStyle.HATCH``, this describes the hatch used.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.FillHatch
        return None

    @fill_hatch.setter
    def fill_hatch(self, value: Hatch) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.FillHatch = value

    @property
    def fill_hatch_name(self) -> str:
        """
        Gets/Sets the name of the used fill hatch style.

        If the property ``fill_style`` is set to ``FillStyle.HATCH``, this is the name of the used fill hatch style.
        """
        return self.__component.FillHatchName

    @fill_hatch_name.setter
    def fill_hatch_name(self, value: str) -> None:
        self.__component.FillHatchName = value

    @property
    def fill_style(self) -> FillStyleProto:
        """
        Gets/Sets the enumeration selects the style the area will be filled with.
        """
        return self.__component.FillStyle

    @fill_style.setter
    def fill_style(self, value: FillStyleProto) -> None:
        self.__component.FillStyle = value

    @property
    def fill_transparence(self) -> int:
        """
        Gets/Sets the transparence of the filled area.

        This property is only valid if the property ``fill_style`` is set to ``FillStyle.SOLID``.
        """
        return self.__component.FillTransparence

    @fill_transparence.setter
    def fill_transparence(self, value: int) -> None:
        self.__component.FillTransparence = value

    @property
    def fill_transparence_gradient(self) -> Gradient | None:
        """
        Gets/Sets the transparency of the fill area as a gradient.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.FillTransparenceGradient
        return None

    @fill_transparence_gradient.setter
    def fill_transparence_gradient(self, value: Gradient) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.FillTransparenceGradient = value

    @property
    def fill_transparence_gradient_name(self) -> str:
        """
        Gets/Sets the name of the used transparence gradient style.

        If a gradient is used for transparency, this is the name of the used transparence gradient style or it is empty.

        If you set the name of a transparence gradient style contained in the document, this style used.
        """
        return self.__component.FillTransparenceGradientName

    @fill_transparence_gradient_name.setter
    def fill_transparence_gradient_name(self, value: str) -> None:
        self.__component.FillTransparenceGradientName = value

    @property
    def fill_use_slide_background(self) -> bool | None:
        """
        If this is ``True``, and ``fill_style`` is ``FillStyle.NONE``: The area displays the slide background.

        **since**

            LibreOffice 7.4

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.FillUseSlideBackground
        return None

    @fill_use_slide_background.setter
    def fill_use_slide_background(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.FillUseSlideBackground = value

    @property
    def graphic_crop(self) -> GraphicCrop | None:
        """
        Gets/Sets the cropping of the object.

        If the property ``fill_bitmap_mode`` is set to ``BitmapMode.STRETCH``, this is the cropping,
        otherwise it is empty.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.GraphicCrop
        return None

    @graphic_crop.setter
    def graphic_crop(self, value: GraphicCrop) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.GraphicCrop = value

    # endregion FillProperties
