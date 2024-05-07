from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.packages.zip import XZipFileAccess2

from ooodev.adapter.component_prop import ComponentProp

from .zip_file_access2_partial import ZipFileAccess2Partial

if TYPE_CHECKING:
    from com.sun.star.packages.zip import ZipFileAccess  # service
    from ooodev.loader.inst.lo_inst import LoInst


class ZipFileAccessComp(ComponentProp, ZipFileAccess2Partial):
    """
    Class for managing ZipFileAccess Component.

    Allows to get reading access to non-encrypted entries inside zip file.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XZipFileAccess2) -> None:
        """
        Constructor

        Args:
            component (XZipFileAccess2): UNO Component that supports ``com.sun.star.packages.zip.ZipFileAccess`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        ZipFileAccess2Partial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.packages.zip.ZipFileAccess",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> ZipFileAccessComp:
        """
        Creates an instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            ZipFileAccessComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XZipFileAccess2, "com.sun.star.packages.zip.ZipFileAccess", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    def create_with_url(self, url: str) -> None:
        """
        Create a new zip file with the given URL.

        Raises:
            com.sun.star.io.IOException: ``IOException``
            com.sun.star.ucb.ContentCreationException: ``ContentCreationException``
            com.sun.star.ucb.InteractiveIOException: ``InteractiveIOException``
            com.sun.star.packages.zip.ZipException: ``ZipException``
        """
        return self.component.createWithURL(url)

    # region Properties
    @property
    def component(self) -> ZipFileAccess:
        """ZipFileAccess Component"""
        # pylint: disable=no-member
        return cast("ZipFileAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
