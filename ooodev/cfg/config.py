# coding: utf-8
from pathlib import Path
from dataclasses import dataclass
from typing import List
import json


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
                # provide defaults because at this time stickytape
                # does not include non *.py files when it packages scripts
                data = {"profile_versions": ["4"]}
            cls._instance = super().__call__(**data)
        return cls._instance


@dataclass(frozen=True)
class Config(metaclass=ConfigMeta):
    """
    Singleton Configuration Class

    This is an internal class and not meant to be used otherwise.

    Never used in macros
    """

    profile_versions: List[str]
    """LibreOffice Profile versions. Currently expect ["4"]"""
