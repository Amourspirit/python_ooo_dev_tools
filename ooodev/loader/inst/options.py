from __future__ import annotations
from dataclasses import dataclass, asdict
import json
import logging


@dataclass(frozen=True)
class Options:
    """
    Lo Load options

    .. versionadded:: 0.6.10
    """

    verbose: bool = False
    """Determines if various info is sent to console. Default ``False``"""

    dynamic: bool = True
    """
    Determines if the script context is dynamic.

    Also When loading a component via :py:meth:`LoInst.load_component() <ooodev.utils.inst.lo.LoInst.load_component>` It is recommended to set this value to ``False``.

    If dynamic the script context is created with the current document; Otherwise, context is static.
    Static context is useful when the script is only going to operate on a a single document.
    There may be a performance gain when using static context.

    Default ``True``
    
    .. versionadded:: 0.11.13
    """
    log_level: int = logging.INFO
    """Logging level. Default ``logging.INFO``"""

    lo_cache_size: int = 200
    """Lo Instance cache size. Default ``200``, ``0`` or less means no caching. Normally you should not need to change this value. If you do, it should be a power of 2."""

    def serialize(self) -> str:
        """
        Serialize the options to a json string.

        Returns:
            str: Json string
        """
        return json.dumps(asdict(self))

    @staticmethod
    def deserialize(s: str) -> Options:
        """
        Deserialize the options from a json string.

        Args:
            s (str): Json string

        Returns:
            Options: Options object
        """
        return Options(**json.loads(s))
