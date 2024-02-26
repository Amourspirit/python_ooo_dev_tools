from __future__ import annotations
from typing import Any, List, TYPE_CHECKING
import contextlib
import uno
from com.sun.star.lang import XServiceInfo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object


class ServicePartialT(Protocol):
    """Type for ServicePartial Class."""

    def support_service(self, *service: str) -> bool:
        """
        Gets if instance supports a service.

        Args:
            *service (str): Variable length argument list of UNO namespace strings such as ``com.sun.star.configuration.GroupAccess``

        Returns:
            bool: ``True`` if instance supports any passed in service; Otherwise, ``False``
        """
        ...

    def get_services(self) -> List[str]:
        """
        Gets service names for the instance.

        Returns:
            List[str]: service names
        """
        ...
