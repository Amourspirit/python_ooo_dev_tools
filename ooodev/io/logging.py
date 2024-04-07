from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import os
import logging
from pprint import pprint
from ooodev.meta.singleton import Singleton


class _Logger(metaclass=Singleton):

    def __init__(self):
        is_windows = os.name == "nt"
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)
        if is_windows:
            logging.addLevelName(logging.ERROR, "ERROR")
            logging.addLevelName(logging.DEBUG, "DEBUG")
            logging.addLevelName(logging.INFO, "INFO")
        else:
            logging.addLevelName(logging.ERROR, "\033[1;41mERROR\033[1;0m")
            logging.addLevelName(logging.DEBUG, "\x1b[33mDEBUG\033[1;0m")
            logging.addLevelName(logging.INFO, "\x1b[32mINFO\033[1;0m")

        log_format = "%(asctime)s - %(levelname)s - %(message)s"
        logging.basicConfig(level=logging.DEBUG, format=log_format, datefmt="%d/%m/%Y %H:%M:%S")
        handler = logging.StreamHandler()
        formatter = logging.Formatter(log_format)
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

    def debug(self, message):
        self.logger.debug(message)

    def info(self, message):
        self.logger.info(message)

    def warning(self, message):
        self.logger.warning(message)

    def error(self, message):
        self.logger.error(message)

    def critical(self, message):
        self.logger.critical(message)


def debug(*messages) -> None:
    """Show messages debug

    :param messages: List of messages to debug
    :type messages: list[Any]
    """
    log = _Logger()
    data = [str(m) for m in messages]
    log.debug("\t".join(data))
    return


def error(message: Any) -> None:
    """Show message error

    :param message: The message error
    :type message: Any
    """
    log = _Logger()
    log.error(message)
    return


def info(*messages) -> None:
    """Show messages info

    :param messages: List of messages to debug
    :type messages: list[Any]
    """
    log = _Logger()
    data = [str(m) for m in messages]
    log.info("\t".join(data))
    return


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
