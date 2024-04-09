from __future__ import annotations
from typing import Any
import os
import logging
from pprint import pprint
from ooodev.meta.singleton import Singleton


class _Logger(metaclass=Singleton):

    def __init__(self):
        is_windows = os.name == "nt"
        self.logger = logging.getLogger(__name__)
        self.logger.propagate = False
        self.logger.setLevel(logging.DEBUG)
        if is_windows:
            logging.addLevelName(logging.ERROR, "ERROR")
            logging.addLevelName(logging.DEBUG, "DEBUG")
            logging.addLevelName(logging.INFO, "INFO")
            logging.addLevelName(logging.WARNING, "WARNING")
        else:
            logging.addLevelName(logging.ERROR, "\033[1;41mERROR\033[1;0m")
            logging.addLevelName(logging.DEBUG, "\x1b[33mDEBUG\033[1;0m")
            logging.addLevelName(logging.INFO, "\x1b[32mINFO\033[1;0m")
            logging.addLevelName(logging.WARNING, "\x1b[32mWARNING\033[1;0m")

        # log_format = "%(asctime)s - %(levelname)s - %(message)s"
        log_format = "{asctime} - {levelname} - {message}"
        logging.basicConfig(level=logging.DEBUG, format=log_format, datefmt="%d/%m/%Y %H:%M:%S", style="{")
        handler = logging.StreamHandler()
        formatter = logging.Formatter(log_format, datefmt="%d/%m/%Y %H:%M:%S", style="{")
        handler.setFormatter(formatter)
        if self.logger.hasHandlers():
            self.logger.handlers.clear()
        self.logger.addHandler(handler)

    def debug(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.debug(msg, *args, **kwargs)

    def info(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.info(msg, *args, **kwargs)

    def warning(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.warning(msg, *args, **kwargs)

    def error(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.error(msg, *args, **kwargs)

    def critical(self, msg: Any, *args: Any, **kwargs: Any) -> None:
        self.logger.critical(msg, *args, **kwargs)

    def get_effective_level(self) -> int:
        return self.logger.getEffectiveLevel()


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
    log.logger.setLevel(level)
    return


def get_log_level() -> int:
    """
    Get the log level.

    Returns:
        int: The log level.
    """
    log = _Logger()
    return log.get_effective_level()


def save_log(path: str, data: Any) -> bool:
    """Save data in file, data append to end and automatic add current time.

    :param path: Path to save log
    :type path: str
    :param data: Data to save in file log
    :type data: Any
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
