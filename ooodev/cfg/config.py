# coding: utf-8
from pathlib import Path
from dataclasses import dataclass
from typing import List
import json
import os


class ConfigMeta(type):
    _instance = None

    def __call__(cls, *args, **kwargs):
        if cls._instance is None:
            root = Path(__file__).parent
            config_file = Path(root, "config.json")
            if config_file.exists():
                with open(config_file, "r") as file:
                    data = json.load(file)
            else:
                # provide defaults because at this time scriptmerge
                # does not include non *.py files when it packages scripts
                data = {"profile_versions": ["4"], "slide_template_path": "share/template/common/presnt/"}

            # get any override values from os.environ
            profile_ver = os.environ.get("OOODEV_CONFIG_PROFILE_VERSION", "")
            if profile_ver:
                data["profile_versions"] = [s.strip() for s in profile_ver.split(",")]
            data["slide_template_path"] = os.environ.get(
                "OOODEV_CONFIG_SLIDE_TEMPLATE_PATH", data["slide_template_path"]
            )

            cls._instance = super().__call__(**data)
        return cls._instance


@dataclass(frozen=True)
class Config(metaclass=ConfigMeta):
    """
    Singleton Configuration Class

    Generally speaking this class is only used internally.
    """

    profile_versions: List[str]
    """
    LibreOffice Profile versions. Currently expect ["4"]

    The value for this property can be set using ``os.environ`` with ``OOODEV_CONFIG_PROFILE_VERSION``.

    ``OOODEV_CONFIG_PROFILE_VERSION`` is a comma separated string such as ``"4"`` or ``"4, 5"``
    """
    slide_template_path: str
    """
    String path such as ``share/template/common/layout/``

    The value for this property can be set using ``os.environ`` with ``OOODEV_CONFIG_SLIDE_TEMPLATE_PATH``.

    ``OOODEV_CONFIG_SLIDE_TEMPLATE_PATH`` is a string such as ``"share/template/common/layout/"``
    """
