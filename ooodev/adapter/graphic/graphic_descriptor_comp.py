from __future__ import annotations
from typing import cast, TYPE_CHECKING
import contextlib
import uno
from ooo.dyn.awt.size import Size

from ooodev.adapter.beans.property_set_comp import PropertySetComp
from ooodev.units.size_mm100 import SizeMM100
from ooodev.units.size_px import SizePX

if TYPE_CHECKING:
    from com.sun.star.graphic import GraphicDescriptor  # service


class GraphicDescriptorComp(PropertySetComp):
    """
    Class for managing GraphicDescriptor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: GraphicDescriptor) -> None:
        """
        Constructor

        Args:
            component (GraphicDescriptor): UNO Component that implements ``com.sun.star.graphic.GraphicDescriptor`` service.
        """
        PropertySetComp.__init__(self, component)  # type: ignore

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.graphic.GraphicDescriptor",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> GraphicDescriptor:
        """GraphicDescriptor Component"""
        # pylint: disable=no-member
        return cast("GraphicDescriptor", self._ComponentBase__get_component())  # type: ignore

    @property
    def alpha(self) -> bool | None:
        """
        Gets/Sets - Indicates that it is a pixel graphic with an alpha channel.

        The status of this flag is not always clear if the graphic was not loaded at all, e.g. in case of just querying for the GraphicDescriptor

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.Alpha
        return None

    @alpha.setter
    def alpha(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.Alpha = value

    @property
    def animated(self) -> bool | None:
        """
        Gets/Sets - Indicates that it is a graphic that consists of several frames that can be played as an animation.

        The status of this flag is not always clear if the graphic was not loaded at all,
        e.g. in case of just querying for the ``GraphicDescriptor``

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.Animated
        return None

    @animated.setter
    def animated(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.Animated = value

    @property
    def bits_per_pixel(self) -> int | None:
        """
        The number of bits per pixel used for the pixel graphic.

        This property is not available for vector graphics and may not be available for some kinds of pixel graphics

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.BitsPerPixel
        return None

    @bits_per_pixel.setter
    def bits_per_pixel(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.component.BitsPerPixel = value

    @property
    def graphic_type(self) -> int:
        """
        Gets/Sets the type of the graphic.
        """
        return self.component.GraphicType

    @graphic_type.setter
    def graphic_type(self, value: int) -> None:
        self.component.GraphicType = value

    @property
    def linked(self) -> bool | None:
        """
        Gets/Sets - Indicates that the graphic is an external linked graphic.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.Linked
        return None

    @linked.setter
    def linked(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.Linked = value

    @property
    def mime_type(self) -> str:
        """
        The MimeType of the loaded graphic.

        The mime can be the original mime type of the graphic source the graphic container was constructed from or it can be the internal mime type image/x-vclgraphic, in which case the original mime type is not available anymore

        Currently, the following mime types are supported for loaded graphics:
        """
        return self.component.MimeType

    @mime_type.setter
    def mime_type(self, value: str) -> None:
        self.component.MimeType = value

    @property
    def origin_url(self) -> str | None:
        """
        Gets/Sets the URL of the location from where the graphic was loaded from.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.OriginURL
        return None

    @origin_url.setter
    def origin_url(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.component.OriginURL = value

    @property
    def size100th_mm(self) -> SizeMM100 | None:
        """
        The Size of the graphic in 100th mm.

        This property may not be available in case of pixel graphics or if the logical size can not be determined correctly for some formats without loading the whole graphic

        **optional**
        """
        with contextlib.suppress(AttributeError):
            sz = self.component.Size100thMM

            return SizeMM100.from_mm100(sz.Width, sz.Height)
        return None

    @size100th_mm.setter
    def size100th_mm(self, value: SizeMM100) -> None:
        with contextlib.suppress(AttributeError):
            sz = Size(value.width.value, value.height.value)
            self.component.Size100thMM = sz

    @property
    def size_pixel(self) -> SizePX | None:
        """
        The Size of the graphic in pixel.

        This property may not be available in case of vector graphics or if the pixel size can not be determined correctly for some formats without loading the whole graphic

        **optional**
        """
        with contextlib.suppress(AttributeError):
            sz = self.component.SizePixel
            return SizePX.from_px(sz.Width, sz.Height)
        return None

    @size_pixel.setter
    def size_pixel(self, value: SizePX) -> None:
        with contextlib.suppress(AttributeError):
            sz = Size(round(value.width.value), round(value.height.value))
            self.component.SizePixel = sz

    @property
    def transparent(self) -> bool | None:
        """
        Indicates that it is a transparent graphic.

        This property is always ``True`` for vector graphics.
        The status of this flag is not always clear if the graphic was not loaded at all,
        e.g. in case of just querying for the ``GraphicDescriptor``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.Transparent
        return None

    @transparent.setter
    def transparent(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.Transparent = value

    # endregion Properties
