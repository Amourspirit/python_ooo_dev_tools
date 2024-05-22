from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.io import XTempFile


from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.io.temp_file_partial import TempFilePartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class TempFileComp(ComponentProp, TempFilePartial):
    """
    Class for managing Pipe Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTempFile) -> None:
        """
        Constructor

        Args:
            component (XTempFile): UNO Component that supports ``com.sun.star.io.TempFile`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        TempFilePartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.io.TempFile",)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TempFileComp:
        """
        Creates an instance from the Lo.

        Creates a new temp file in the LibreOffice Temp directory.
        ``resource_name`` something like ``'/tmp/lu23986334kdxgl.tmp/lu23986334kdxhw.tmp'``.
        ``uri`` something like ``'file:///tmp/lu23986334kdxgl.tmp/lu23986334kdxhw.tmp'``.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TempFileComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XTempFile, "com.sun.star.io.TempFile", raise_err=True)  # type: ignore
        return cls(inst)

    @staticmethod
    def context_manager():
        """
        Context manager for managing TempFileComp.

        Returns:
            _TempFileContextManager: Context instance.

        Example:
            .. code-block:: python

                with TempFileComp.context_manager() as tmp:
                    print(tmp.resource_name)
                    print(tmp.uri)
        """
        return _TempFileContextManager()

    # endregion Static Methods


class _TempFileContextManager:
    """
    Context Manager for managing TempFileComp.

    Example:
    ```python
    with TempFileContextManager() as tmp:
        print(tmp.resource_name)
        print(tmp.uri)
    ```
    """

    def __init__(self) -> None:
        self._comp = None

    def __enter__(self) -> TempFileComp:
        self._comp = TempFileComp.from_lo()
        self._comp.remove_file = True
        return self._comp

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self._comp = None
