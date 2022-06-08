# coding: utf-8
import platform
from enum import Enum

class SysInfo:
    """System Info related"""
    class PlatformEnum(str, Enum):
        UNKNOWN = 'Unknown'
        WINDOWS = "Windows"
        MAC = "Darwin"
        LINUX ="Linux"
    
    @staticmethod
    def get_platform() -> 'SysInfo.PlatformEnum':
        """
        Gets current platform

        Returns:
            SysInfo.PlatformEnum: Platform Enum
        """
        s = platform.system().lower()
        if s == 'windows':
            return SysInfo.PlatformEnum.WINDOWS
        if s == 'darwin':
            return SysInfo.PlatformEnum.MAC
        if s == 'linux':
            return SysInfo.PlatformEnum.LINUX
        return SysInfo.PlatformEnum.UNKNOWN
        