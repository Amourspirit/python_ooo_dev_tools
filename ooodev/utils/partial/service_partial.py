from __future__ import annotations
from typing import Any, List
import contextlib
import uno
from com.sun.star.lang import XServiceInfo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.loader import lo as mLo


class ServicePartial:
    """Partial Class used for getting service information from a UNO component."""

    def __init__(self, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor.

        Args:
            component (Any): Any Uno Component that supports ``XServiceInfo`` interface.
            lo_inst (LoInst, optional): Lo instance.
        """
        if lo_inst is None:
            self.__lo_inst = mLo.Lo.current_lo
        else:
            self.__lo_inst = lo_inst
        self.__component = component

    def support_service(self, *service: str) -> bool:
        """
        Gets if instance supports a service.

        Args:
            *service (str): Variable length argument list of UNO namespace strings such as ``com.sun.star.configuration.GroupAccess``

        Returns:
            bool: ``True`` if instance supports any passed in service; Otherwise, ``False``
        """

        result = False
        with contextlib.suppress(Exception):
            si = self.__lo_inst.qi(XServiceInfo, self.__component)
            if si is None:
                return result
            for srv in service:
                result = si.supportsService(srv)
                if result:
                    break
        return result

    def get_services(self) -> List[str]:
        """
        Gets service names for the instance.

        Returns:
            List[str]: service names
        """
        with contextlib.suppress(AttributeError):
            return sorted(self.__component.SupportedServiceNames)  # type: ignore
        si = self.__lo_inst.qi(XServiceInfo, self.__component, True)
        names = si.getSupportedServiceNames()
        return sorted(names)
