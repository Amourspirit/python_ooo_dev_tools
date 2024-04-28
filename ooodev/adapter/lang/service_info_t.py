from __future__ import annotations
from typing import Tuple, Protocol
import uno


class ServiceInfoT(Protocol):
    """
    Protocol class for ``XServiceInfo``.
    """

    # region XServiceInfo
    def getImplementationName(self) -> str:
        """
        Provides the implementation name of the service implementation.
        """
        ...

    def getSupportedServiceNames(self) -> Tuple[str, ...]:
        """
        Provides the supported service names of the implementation, including also indirect service names.
        """
        ...

    def supportsService(self, name: str) -> bool:
        """
        Tests whether any of the specified service(s) are supported.

        Args:
            name (str): Service name to test such as ``com.sun.star.awt.MenuBar``.

        Returns:
            bool: ``True`` if specified service is supported; Otherwise, ``False``.
        """
        ...

    # endregion XServiceInfo
