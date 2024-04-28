from __future__ import annotations
from typing import Tuple
import uno
from com.sun.star.lang import XServiceInfo


class BaseServiceInfo(XServiceInfo):
    """
    Provides information regarding the implementation: which services are implemented and the name of the implementation.

    See Also:
        `API XServiceInfo <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XServiceInfo.html>`_
    """

    # region XServiceInfo
    def getImplementationName(self) -> str:
        """
        Provides the implementation name of the service implementation.
        """
        raise NotImplementedError

    def supportsService(self, name: str) -> bool:
        """
        Tests whether the specified service is supported, i.e.

        implemented by the implementation.
        """
        raise NotImplementedError

    def getSupportedServiceNames(self) -> Tuple[str]:
        """
        Provides the supported service names of the implementation, including also indirect service names.
        """
        raise NotImplementedError

    # endregion XServiceInfo
