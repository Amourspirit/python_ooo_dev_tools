from __future__ import annotations
from typing import Any
from ooodev.io.log import logging as logger
import logging


class NamedLogger:
    """
    Class for Logging class messages
    """

    def __init__(self, name: str) -> None:
        """
        Constructor

        Args:
            name (str): Name of the class.
        """
        self._name = name
        self._logging_level = logger.get_log_level()

    def debug(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        """
        Logs debug message.

        Args:
            msg (Any):  message to debug.
            args (Any, optional):  arguments.

        Keyword Args:
            exc_info: (_ExcInfoType): Exc Info Type Default to  ``None``
            stack_info (bool): Stack Info. Defaults to  ``False``.
            stacklevel (int): Stack Level. Defaults to  ``1``.
            extra (Mapping[str, object], None): extra Defaults to ``None``.

        Returns:
            None
        """
        logger.debug(f"{self._name}: {msg}", *args, **kwargs)
        return

    def debugs(self, *messages) -> None:
        """
        Show messages debug

        Args:
            messages (list[Any]): List of messages to debug
        """
        data = [str(m) for m in messages]
        data.insert(0, f"{self._name}:")
        logger.debugs(*data)
        return

    def error(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        """
        Logs error message.

        Args:
            msg (Any):  message to debug.
            args (Any, optional):  arguments.

        Keyword Args:
            exc_info: (_ExcInfoType): Exc Info Type Default to  ``None``
            stack_info (bool): Stack Info. Defaults to  ``False``.
            stacklevel (int): Stack Level. Defaults to  ``1``.
            extra (Mapping[str, object], None): extra Defaults to ``None``.

        Returns:
            None
        """
        logger.error(f"{self._name}: {msg}", *args, **kwargs)
        return

    def info(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        """
        Logs info message.

        Args:
            msg (Any):  message to debug.
            args (Any, optional):  arguments.

        Keyword Args:
            exc_info: (_ExcInfoType): Exc Info Type Default to  ``None``
            stack_info (bool): Stack Info. Defaults to  ``False``.
            stacklevel (int): Stack Level. Defaults to  ``1``.
            extra (Mapping[str, object], None): extra Defaults to ``None``.

        Returns:
            None
        """
        logger.info(f"{self._name}: {msg}", *args, **kwargs)
        return

    def warning(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        """
        Logs info message.

        Args:
            msg (Any):  message to debug.
            args (Any, optional):  arguments.

        Keyword Args:
            exc_info: (_ExcInfoType): Exc Info Type Default to  ``None``
            stack_info (bool): Stack Info. Defaults to  ``False``.
            stacklevel (int): Stack Level. Defaults to  ``1``.
            extra (Mapping[str, object], None): extra Defaults to ``None``.

        Returns:
            None
        """
        logger.warning(f"{self._name}: {msg}", *args, **kwargs)
        return

    # region Properties
    @property
    def is_debug(self) -> bool:
        """Check if is debug"""
        return self._logging_level >= logging.DEBUG

    @property
    def is_info(self) -> bool:
        """Check if is info"""
        return self._logging_level >= logging.INFO

    @property
    def is_warning(self) -> bool:
        """Check if is warning"""
        return self._logging_level >= logging.WARNING

    @property
    def is_error(self) -> bool:
        """Check if is error"""
        return self._logging_level >= logging.ERROR

    # endregion Properties
