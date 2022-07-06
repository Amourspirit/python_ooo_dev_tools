# coding: utf-8
from __future__ import annotations
import uno
import unohelper
from typing import Tuple
from com.sun.star.datatransfer import XTransferable
from ooo.dyn.datatransfer.unsupported_flavor_exception import UnsupportedFlavorException
from ooo.dyn.datatransfer.data_flavor import DataFlavor


class TextTransferable(unohelper.Base, XTransferable):

    MIME_TYPE = "text/plain;charset=utf-16"

    def __init__(self) -> None:
        super().__init__()
        self._im_bytes = []

    def getTransferData(self, aFlavor: DataFlavor) -> object:
        """
        Called by a data consumer to obtain data from the source in a specified format.

        Args:
            aFlavor (DataFlavor): DataFlavor struct

        Raises:
            UnsupportedFlavorException: ``UnsupportedFlavorException``
            com.sun.star.io.IOException: ``IOException``
        """
        if aFlavor.MimeType == TextTransferable.MIME_TYPE:
            return self._im_bytes
        raise UnsupportedFlavorException

    def getTransferDataFlavors(self) -> Tuple[DataFlavor, ...]:
        """
        Returns a sequence of supported DataFlavor.
        """
        # not yet sure the proper DataType.
        # In java the DataType is Type(byte[].class)
        dfs = (DataFlavor(MimeType=TextTransferable.MIME_TYPE, HumanPresentableName="Bitmap", DataType=tuple),)
        return dfs

    def isDataFlavorSupported(self, aFlavor: DataFlavor) -> bool:
        """
        Checks if the data object supports the specified by :py:attr:`.TextTransferable.MIME_TYPE`.

        A value of False if the DataFlavor is unsupported by the transfer source.

        Args:
            aFlavor (DataFlavor): DataFlavor struct

        Note:
            This method is only for analogy with the JAVA Clipboard interface.
            To avoid many calls, the caller should instead use
            ``com.sun.star.datatransfer.XTransferable.getTransferDataFlavors()``.
        """
        return aFlavor.MimeType == TextTransferable.MIME_TYPE
