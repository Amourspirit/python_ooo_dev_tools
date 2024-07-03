from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import os
import sys
import logging
from logging.handlers import TimedRotatingFileHandler
from pprint import pprint

if TYPE_CHECKING:
    from ooodev.utils.type_var import PathOrStr


class _Logger:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(_Logger, cls).__new__(cls)
            cls._instance._is_init = False
        return cls._instance

    def __init__(self):
        if getattr(self, "_is_init", False):
            return
        # is_windows = os.name == "nt"
        try:
            self._log_level = int(os.environ.get("ODEV_LOG_LEVEL", logging.INFO))
        except Exception:
            self._log_level = logging.INFO
        self._file_handlers: Dict[str, TimedRotatingFileHandler] = {}

        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(self._log_level)
        self.logger.propagate = False

        logging.addLevelName(logging.ERROR, "ERROR")
        logging.addLevelName(logging.DEBUG, "DEBUG")
        logging.addLevelName(logging.INFO, "INFO")
        logging.addLevelName(logging.WARNING, "WARNING")
        logging.addLevelName(logging.CRITICAL, "CRITICAL")
        # if is_windows:
        #     logging.addLevelName(logging.ERROR, "ERROR")
        #     logging.addLevelName(logging.DEBUG, "DEBUG")
        #     logging.addLevelName(logging.INFO, "INFO")
        #     logging.addLevelName(logging.WARNING, "WARNING")
        # else:
        #     logging.addLevelName(logging.ERROR, "\033[1;41mERROR\033[1;0m")
        #     logging.addLevelName(logging.DEBUG, "\x1b[33mDEBUG\033[1;0m")
        #     logging.addLevelName(logging.INFO, "\x1b[32mINFO\033[1;0m")
        #     logging.addLevelName(logging.WARNING, "\x1b[32mWARNING\033[1;0m")

        # log_format = "%(asctime)s - %(levelname)s - %(message)s"
        log_format = "{asctime} - {levelname} - {message}"
        logging.basicConfig(level=logging.DEBUG, format=log_format, datefmt="%d/%m/%Y %H:%M:%S", style="{")

        self._formatter = logging.Formatter(log_format, datefmt="%d/%m/%Y %H:%M:%S", style="{")
        if self.logger.hasHandlers():
            self.logger.handlers.clear()
        if self._log_level >= logging.DEBUG:
            stream_handler = self._get_console_handler()
            self.logger.addHandler(stream_handler)
        else:
            self.logger.addHandler(self._get_null_handler())
        self._is_init = True

    def debug(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.debug(msg, *args, **kwargs)

    def info(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.info(msg, *args, **kwargs)

    def warning(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.warning(msg, *args, **kwargs)

    def error(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.error(msg, *args, **kwargs)

    def exception(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.exception(msg, *args, **kwargs)

    def critical(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.critical(msg, *args, **kwargs)

    def get_effective_level(self) -> int:
        return self.logger.getEffectiveLevel()

    # region handlers

    def add_console_handler(self):
        if not self.is_stream_handler:
            self.logger.addHandler(self._get_console_handler())

    def _get_console_handler(self):
        # check to see if there is already a console handler
        for handler in self.logger.handlers:
            if isinstance(handler, logging.StreamHandler):
                return handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(self._formatter)
        console_handler.setLevel(self._log_level)
        return console_handler

    def _get_null_handler(self):
        return logging.NullHandler()

    def remove_handlers(self):
        """
        Remove all handlers from the logger.
        """
        self.logger.handlers.clear()
        self.logger.addHandler(self._get_null_handler())

    def _get_file_handler(self, log_file: PathOrStr, log_level: int = -1):
        key = str(log_file)
        if key in self._file_handlers:
            file_handler = self._file_handlers[key]
        else:
            file_handler = TimedRotatingFileHandler(
                log_file, when="W0", interval=1, backupCount=3, encoding="utf8", delay=True
            )
            # file_handler = logging.FileHandler(log_file, mode="w", encoding="utf8", delay=True)
            file_handler.setFormatter(self._formatter)
            self._file_handlers[key] = file_handler
        if log_level != -1:
            file_handler.setLevel(log_level)
        else:
            file_handler.setLevel(self._log_level)
        return file_handler

    def has_file_handler(self, log_file: PathOrStr):
        key = str(log_file)
        return key in self._file_handlers

    def add_file_handler(self, log_file: PathOrStr, log_level: int = -1) -> bool:
        """
        Add a file handler to the logger if it does not already exist.

        Args:
            log_file (PathOrStr): Log File Path.
            log_level (int, optional): Log Level. Defaults to Instance Log Level.

        Returns:
            bool: True if the handler was added, False otherwise.
        """
        handler = self._get_file_handler(log_file, log_level)
        if handler not in self.logger.handlers:
            self.logger.addHandler(handler)
            return True
        return False

    def remove_file_handler(self, log_file: PathOrStr) -> bool:
        """
        Remove a file handler from the logger if it exists.

        Args:
            log_file (PathOrStr): Log File Path.

        Returns:
            bool: True if the handler was removed, False otherwise.
        """
        key = str(log_file)
        if key in self._file_handlers:
            handler = self._file_handlers[key]
            self.logger.removeHandler(handler)
            del self._file_handlers[key]
            return True
        return False

    def add_stream_handler(self):
        """Adds a stream handler to the logger if it does not already exist."""
        # only add stream handler if it does nto already exist.
        for handler in self.logger.handlers:
            if isinstance(handler, logging.StreamHandler):
                return
        self.logger.addHandler(self._get_console_handler())

    @property
    def is_stream_handler(self) -> bool:
        """Check if logger has a stream handler"""
        for handler in self.logger.handlers:
            if isinstance(handler, logging.StreamHandler):
                return True
        return False

    @property
    def is_file_handler(self) -> bool:
        """Check if logger has a file handler"""
        for handler in self.logger.handlers:
            if isinstance(handler, logging.FileHandler):
                return True
        return False

    @property
    def log_level(self) -> int:
        return self._log_level

    @log_level.setter
    def log_level(self, value: int) -> None:
        self._log_level = value
        self.logger.setLevel(value)
        if value > 0:
            if self.is_file_handler is False and self.is_stream_handler is False:
                self.add_stream_handler()
        else:
            self.remove_handlers()  # add null handler only
        for handler in self.logger.handlers:
            handler.setLevel(value)
        return

    # endregion handlers

    @classmethod
    def reset_instance(cls):
        cls._instance = None


def debug(msg: Any, *args: Any, **kwargs: Any) -> None:
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
    log = _Logger()
    log.debug(msg, *args, **kwargs)
    return


def critical(msg: Any, *args: Any, **kwargs: Any) -> None:
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
    log = _Logger()
    log.critical(msg, *args, **kwargs)
    return


def debugs(*messages: str) -> None:
    """
    Log Several messages debug formatted by tab.

    Args:
        messages (Any):  One or more messages to log.

    Return:
        None:
    """
    data = [str(m) for m in messages]
    debug("\t".join(data))
    return


def error(msg: Any, *args: Any, **kwargs: Any) -> None:
    """
    Logs Error message.

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
    log = _Logger()
    log.error(msg, *args, **kwargs)
    return


def exception(msg: Any, *args: Any, **kwargs: Any) -> None:
    """
    Logs exception message.

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
    log = _Logger()
    log.exception(msg, *args, **kwargs)
    return


def info(msg: Any, *args: Any, **kwargs: Any) -> None:
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
    log = _Logger()
    log.info(msg, *args, **kwargs)
    return


def warning(msg: Any, *args: Any, **kwargs: Any) -> None:
    """
    Logs warning message.

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
    log = _Logger()
    log.warning(msg, *args, **kwargs)
    return


def infos(*messages: Any) -> None:
    """
    Log Several messages info formatted by tab.

    Args:
        messages (Any):  One or more messages to log.

    Return:
        None:
    """
    log = _Logger()
    data = [str(m) for m in messages]
    log.info("\t".join(data))
    return


def set_log_level(level: int) -> None:
    """
    Set the log level.

    Args:
        level (int): The log level.
    """
    log = _Logger()
    log.log_level = level
    return


def get_log_level() -> int:
    """
    Get the log level.

    Returns:
        int: The log level.
    """
    log = _Logger()
    return log.get_effective_level()


def save_log(path: Any, data: Any) -> bool:
    """Save data in file, data append to end and automatic add current time.

    Args:
        path (Any): PathLike path such a string or Path to save log.
        data (Any): Data to save in file log

    Returns:
        bool: True if success, False otherwise
    """
    # DateUtil import logging
    from ooodev.utils.date_time_util import DateUtil

    result = True

    try:
        with open(path, "a") as f:
            f.write(f"{str(DateUtil.now)} - ")
            pprint(data, stream=f)
    except Exception as e:
        error(e)
        result = False

    return result


def add_file_logger(log_file: PathOrStr, log_level: int = -1) -> bool:
    """
    Add a file logger to the logger if it does not already exist.

    Args:
        log_file (PathOrStr): Log File Path.
        log_level (int, optional): Log Level. Defaults to Instance Log Level.

    Returns:
        bool: True if the handler was added, False otherwise.
    """
    log = _Logger()
    return log.add_file_handler(log_file, log_level)


def remove_file_logger(log_file: PathOrStr) -> bool:
    """
    Remove a file logger from the logger if it exists.

    Args:
        log_file (PathOrStr): Log File Path.

    Returns:
        bool: True if the handler was removed, False otherwise.
    """
    log = _Logger()
    return log.remove_file_handler(log_file)


def remove_handlers() -> None:
    """
    Remove all handlers from the logger.
    """
    log = _Logger()
    log.remove_handlers()
    return


def add_stream_handler() -> None:
    """Adds a stream handler to the logger if it does not already exist."""
    log = _Logger()
    log.add_stream_handler()
    return


def reset_logger():
    """Reset the logger instance."""
    _Logger.reset_instance()
    return
