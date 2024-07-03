from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.io.log import logging as logger
import logging

if TYPE_CHECKING:
    from ooodev.utils.type_var import PathOrStr


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

    def exception(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        """
        Logs error message.

        Args:
            msg (Any):  message to debug.
            args (Any, optional):  arguments.

        Keyword Args:
            exc_info: (_ExcInfoType): Exc Info Type Default to  ``True``
            stack_info (bool): Stack Info. Defaults to  ``False``.
            stacklevel (int): Stack Level. Defaults to  ``1``.
            extra (Mapping[str, object], None): extra Defaults to ``None``.

        Returns:
            None
        """
        logger.exception(f"{self._name}: {msg}", *args, **kwargs)
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

    def critical(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        """
        Logs critical message.

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
        logger.critical(f"{self._name}: {msg}", *args, **kwargs)
        return

    # region Handler methods
    def add_file_logger(self, log_file: PathOrStr, log_level: int = -1) -> bool:
        """
        Add a file logger to the logger if it does not already exist.

        Args:
            log_file (PathOrStr): Log File Path.
            log_level (int, optional): Log Level. Defaults to Instance Log Level.

        Returns:
            bool: True if the handler was added, False otherwise.
        """
        return logger.add_file_logger(log_file, log_level)

    def remove_file_logger(self, log_file: PathOrStr) -> bool:
        """
        Remove a file logger from the logger if it exists.

        Args:
            log_file (PathOrStr): Log File Path.

        Returns:
            bool: True if the handler was removed, False otherwise.
        """
        return logger.remove_file_logger(log_file)

    def remove_handlers(self) -> None:
        """
        Remove all handlers from the logger.
        """
        logger.remove_handlers()

    def add_stream_handler(self) -> None:
        """Adds a stream handler to the logger if it does not already exist."""
        logger.add_stream_handler()

    # endregion Handler methods

    # region Properties
    @property
    def is_debug(self) -> bool:
        """Check if is debug"""
        return self._logging_level <= logging.DEBUG

    @property
    def is_info(self) -> bool:
        """Check if is info"""
        return self._logging_level <= logging.INFO

    @property
    def is_warning(self) -> bool:
        """Check if is warning"""
        return self._logging_level <= logging.WARNING

    @property
    def is_error(self) -> bool:
        """Check if is error"""
        return self._logging_level <= logging.ERROR

    # endregion Properties
