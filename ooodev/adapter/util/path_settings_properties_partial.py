from __future__ import annotations
import contextlib
from typing import Any, Tuple
from pathlib import Path
import uno
from ooodev.utils.string.str_list import StrList


class PathSettingsPropertiesPartial:
    """
    Path Properties Partial Class.

    See Also:
        `API XPathSettings <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XPathSettings.html>`_
    """

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.util.XPathSettings`` interface.
        """
        self.__component = component

    # region path methods
    def _process_path(self, path: str) -> str:
        """
        Process the path.

        Args:
            path (str): Path to process.

        Returns:
            str: Processed path.
        """
        if path.startswith("file://"):
            return path
        return uno.systemPathToFileUrl(path)

    def _process_single_path(self, value: str | Path | StrList | StrList) -> str:
        if isinstance(value, StrList):
            s = value[0]
        else:
            s = str(value)
        return self._process_path(s)

    def _process_paths(self, path: str) -> str:
        """
        Process the paths that are separated by a semicolon.

        Args:
            paths (str): Paths to process.

        Returns:
            str: Processed paths.
        """
        # split the paths by semicolon
        split_paths = path.split(";")
        # process each path
        paths = [self._process_path(p) for p in split_paths]
        return ";".join(paths)

    def _process_multi_path(self, value: str | Path | StrList | Tuple[Path, ...] | StrList | StrList) -> str:
        if isinstance(value, tuple):
            s = ";".join([str(v) for v in value])
        else:
            s = str(value)
        return self._process_paths(s)

    # endregion path methods

    # region XPathSettings
    @property
    def addin(self) -> StrList:
        """
        Gets/Sets the directory that contains spreadsheet add-ins which use the old add-in API.

        Contains spreadsheet add-ins that use the old add-in API.

        When setting can be a str, StrList, or a path like object.

        Support single path only.
        """
        return StrList.from_str(self.__component.Addin)

    @addin.setter
    def addin(self, value: str | Path | StrList | StrList) -> None:
        self.__component.Addin = self._process_single_path(value)

    @property
    def auto_correct(self) -> StrList:
        """
        Gets/Sets the settings of the auto_correct dialog.

        Contains the settings for the AutoCorrect dialog.

        The value can be more than one path separated by a semicolon.

        When setting can be a str or a path like object, StrList or a Tuple of path like object.

        Supports multiple paths.
        """
        return StrList.from_str(self.__component.AutoCorrect)

    @auto_correct.setter
    def auto_correct(self, value: str | Path | StrList | Tuple[Path, ...] | StrList) -> None:
        self.__component.AutoCorrect = self._process_multi_path(value)

    @property
    def auto_text(self) -> StrList:
        """
        Gets/Sets the directory which contains the AutoText modules.

        Contains the AutoText modules.

        The value can be more than one path separated by a semicolon.

        When setting can be a str or a path like object, StrList or a Tuple of path like object.

        Supports multiple paths.
        """
        return StrList.from_str(self.__component.AutoText)

    @auto_text.setter
    def auto_text(self, value: str | Path | StrList | Tuple[Path, ...] | StrList) -> None:
        self.__component.AutoText = self._process_multi_path(value)

    @property
    def backup(self) -> StrList:
        """
        Gets/Sets the automatic backup copies of documents are stored here.

        Where automatic document backups are stored.

        When setting can be a str, StrList, or a path like object.

        Support single path only.
        """
        return StrList.from_str(self.__component.Backup)

    @backup.setter
    def backup(self, value: str | Path | StrList) -> None:
        self.__component.Backup = self._process_single_path(value)

    @property
    def base_path_share_layer(self) -> StrList:
        """
        Gets/Sets the base path for shared layers.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.BasePathShareLayer)

    @base_path_share_layer.setter
    def base_path_share_layer(self, value: str | Path | StrList) -> None:
        self.__component.BasePathShareLayer = self._process_single_path(value)

    @property
    def base_path_user_layer(self) -> StrList:
        """
        Gets/Sets the base path for user layers.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.BasePathUserLayer)

    @base_path_user_layer.setter
    def base_path_user_layer(self, value: str | Path | StrList) -> None:
        self.__component.BasePathUserLayer = self._process_single_path(value)

    @property
    def basic(self) -> StrList:
        """
        Gets/Sets the Basic files, used by the AutoPilots, can be found here.

        Where automatic document backups are stored.

        The value can be more than one path separated by a semicolon or a tuple of path like objects.

        When setting can be a str or a path like object, StrList or a Tuple of path like object.

        Supports multiple paths.
        """
        return StrList.from_str(self.__component.Basic)

    @basic.setter
    def basic(self, value: str | Path | StrList | Tuple[Path, ...] | StrList) -> None:
        self.__component.Basic = self._process_multi_path(value)

    @property
    def bitmap(self) -> StrList:
        """
        Gets/Sets - the directory contains the icons for the toolbars.

        Contains the external icons for the toolbars.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.Bitmap)

    @bitmap.setter
    def bitmap(self, value: str | Path | StrList) -> None:
        self.__component.Bitmap = self._process_single_path(value)

    @property
    def classification(self) -> StrList:
        """
        Gets/Sets - Not documented.


        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        with contextlib.suppress(Exception):
            return StrList.from_str(self.__component.Classification)
        return StrList()

    @classification.setter
    def classification(self, value: str | Path | StrList) -> None:
        self.__component.Classification = self._process_single_path(value)

    @property
    def config(self) -> StrList:
        """
        Gets configuration files are located here.

        Contains configuration files. This property is not visible in the path options dialog and cannot be modified.

        Supports single path only.
        """
        return StrList.from_str(self.__component.Config)

    @property
    def dictionary(self) -> StrList:
        """
        Gets/Sets provided dictionaries are stored here.

        Contains the dictionaries.

        When setting can be a str, StrList, or a path like object.

        Support single path only.
        """
        return StrList.from_str(self.__component.Dictionary)

    @dictionary.setter
    def dictionary(self, value: str | Path | StrList) -> None:
        self.__component.Dictionary = self._process_single_path(value)

    @property
    def favorite(self) -> StrList:
        """
        Gets/Sets Path to save folder bookmarks.

        Contains the saved folder bookmarks.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.Favorite)

    @favorite.setter
    def favorite(self, value: str | Path | StrList) -> None:
        self.__component.Favorite = self._process_single_path(value)

    @property
    def filter(self) -> StrList:
        """
        Gets/Sets the directory where all the filters are stored.

        Where the filters are stored.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.Filter)

    @filter.setter
    def filter(self, value: str | Path | StrList) -> None:
        self.__component.Filter = self._process_single_path(value)

    @property
    def fingerprint(self) -> StrList:
        """
        Gets/Sets Not documented.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        with contextlib.suppress(Exception):
            return StrList.from_str(self.__component.Fingerprint)
        return StrList()

    @fingerprint.setter
    def fingerprint(self, value: str | Path | StrList) -> None:
        self.__component.Fingerprint = self._process_single_path(value)

    @property
    def gallery(self) -> StrList:
        """
        Gets/Sets the directories which contains the Gallery database and multimedia files.

        Contains the Gallery database and multimedia files.

        The value can be more than one path separated by a semicolon or a tuple of path like objects.

        When setting can be a str or a path like object, StrList or a Tuple of path like object.

        Supports multiple paths.
        """
        return StrList.from_str(self.__component.Gallery)

    @gallery.setter
    def gallery(self, value: str | Path | StrList | Tuple[Path, ...] | StrList) -> None:
        self.__component.Gallery = self._process_multi_path(value)

    @property
    def graphic(self) -> StrList:
        """
        Gets/Sets - This directory is displayed when the dialog for opening a graphic or for saving a new graphic is called.

        When setting can be a str, StrList, or a path like object.
        """
        return StrList.from_str(self.__component.Graphic)

    @graphic.setter
    def graphic(self, value: str | Path | StrList) -> None:
        self.__component.Graphic = self._process_single_path(value)

    @property
    def help(self) -> StrList:
        """
        Gets/Sets the path to the Office help files.

        Contains the OOo help files.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.Help)

    @help.setter
    def help(self, value: str | Path | StrList) -> None:
        self.__component.Help = self._process_single_path(value)

    @property
    def iconset(self) -> StrList:
        """
        Gets/Sets Not documented.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        with contextlib.suppress(Exception):
            return StrList.from_str(self.__component.Iconset)
        return StrList()

    @iconset.setter
    def iconset(self, value: str | Path | StrList) -> None:
        self.__component.Iconset = self._process_single_path(value)

    @property
    def linguistic(self) -> StrList:
        """
        Gets/Sets the files that are necessary for the spell check are saved here.

        Contains the OOo spellcheck files.

        The value can be more than one path separated by a semicolon or a tuple of path like objects.

        When setting can be a str or a path like object, StrList or a Tuple of path like object.

        Supports multiple paths.
        """
        return StrList.from_str(self.__component.Linguistic)

    @linguistic.setter
    def linguistic(self, value: str | Path | StrList | Tuple[Path, ...] | StrList) -> None:
        self.__component.Linguistic = self._process_multi_path(value)

    @property
    def module(self) -> StrList:
        """
        Gets/Sets the path for the modules.

        Contains the OOo modules.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.Module)

    @module.setter
    def module(self, value: str | Path | StrList) -> None:
        self.__component.Module = self._process_single_path(value)

    @property
    def palette(self) -> StrList:
        """
        Gets/Sets the path to the palette files ``*.SOB`` to ``*.SOF`` containing user-defined colors and patterns.

        Contains the palette files that contain user-defined colors and patterns (``*.SOB`` and ``*.SOF``).

        The value can be more than one path separated by a semicolon or a tuple of path like objects.

        When setting can be a str or a path like object, StrList or a Tuple of path like object.

        Supports multiple paths.
        """
        return StrList.from_str(self.__component.Palette)

    @palette.setter
    def palette(self, value: str | Path | StrList | Tuple[Path, ...] | StrList) -> None:
        self.__component.Palette = self._process_multi_path(value)

    @property
    def plugin(self) -> StrList:
        """
        Gets/Sets - Plugins are saved in these directories.

        The value can be more than one path separated by a semicolon or a tuple of path like objects.

        When setting can be a str or a path like object, StrList or a Tuple of path like object.

        Supports multiple paths.
        """
        return StrList.from_str(self.__component.Plugin)

    @plugin.setter
    def plugin(self, value: str | Path | StrList | Tuple[Path, ...] | StrList) -> None:
        self.__component.Plugin = self._process_multi_path(value)

    @property
    def storage(self) -> StrList:
        """
        Gets/Sets - Mail, News files and other information (for example, about FTP Server) are stored here.

        When setting can be a str, StrList, or a path like object.

        Support single path only.
        """
        with contextlib.suppress(Exception):
            return StrList.from_str(self.__component.Storage)
        return StrList()

    @storage.setter
    def storage(self, value: str | Path | StrList) -> None:
        self.__component.Storage = self._process_single_path(value)

    @property
    def temp(self) -> StrList:
        """
        Gets/Sets the base url to the office temp-files.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.Temp)

    @temp.setter
    def temp(self, value: str | Path | StrList) -> None:
        self.__component.Temp = self._process_single_path(value)

    @property
    def template(self) -> StrList:
        """
        The templates originate from these folders and sub-folders.

        The value can be more than one path separated by a semicolon or a tuple of path like objects.

        When setting can be a str or a path like object, StrList or a Tuple of path like object.

        Supports multiple paths.
        """
        return StrList.from_str(self.__component.Template)

    @template.setter
    def template(self, value: str | Path | StrList | Tuple[Path, ...] | StrList) -> None:
        self.__component.Template = self._process_multi_path(value)

    @property
    def ui_config(self) -> StrList:
        """
        Global directories to look for user interface configuration files.

        The user interface configuration will be merged with user settings stored in the directory specified by UserConfig.
        The value can be more than one path separated by a semicolon or a tuple of path like objects.

        When setting can be a str or a path like object, StrList or a Tuple of path like object.

        Supports multiple paths.
        """
        return StrList.from_str(self.__component.UIConfig)

    @ui_config.setter
    def ui_config(self, value: str | Path | StrList | Tuple[Path, ...] | StrList) -> None:
        self.__component.UIConfig = self._process_multi_path(value)

    @property
    def user_config(self) -> StrList:
        """
        Specifies the folder with the user settings.

        Contains the user settings, including the user interface configuration files for menus, toolbars, accelerators and status bars.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.UserConfig)

    @user_config.setter
    def user_config(self, value: str | Path | StrList) -> None:
        self.__component.UserConfig = self._process_single_path(value)

    @property
    def user_dictionary(self) -> StrList:
        """
        The custom dictionaries are contained here.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.UserDictionary)

    @user_dictionary.setter
    def user_dictionary(self, value: str | Path | StrList) -> None:
        self.__component.UserDictionary = self._process_single_path(value)

    @property
    def work(self) -> StrList:
        """
        The path of the work folder can be modified according to the user's needs.

        The work folder. This path can be modified according to the user's needs and can be seen in the Open or Save dialog.

        The path specified here can be seen in the Open or Save dialog.

        When setting can be a str, StrList, or a path like object.

        Supports single path only.
        """
        return StrList.from_str(self.__component.Work)

    @work.setter
    def work(self, value: str | Path | StrList) -> None:
        self.__component.Work = self._process_single_path(value)

    # endregion XPathSettings
