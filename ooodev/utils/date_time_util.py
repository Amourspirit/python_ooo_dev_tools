# coding: utf-8
from __future__ import annotations
import datetime
import numbers
from typing import cast
from . import lo as mLo
from com.sun.star.util import DateTime as UnoDateTime


class DateUtil:
    """Date and time utilities"""

    # see also:Talk - Benjamin "Zags" Zagorsky: Handling Timezones in Python
    # https://www.youtube.com/watch?v=XZlPXLsSU2U&t=1283s

    @staticmethod
    def time_stamp(tz: datetime.timezone | None = None) -> str:
        """
        Gets a time stamp string

        Args:
            tz (timezone | None, optional): TimeZone

        Returns:
            str: Formatted timestamp such as ``2022-06-19 17:12:38``
        """
        if tz is not None:
            dt = datetime.datetime.now(tz)
        else:
            dt = datetime.datetime.now()
        return dt.strftime("%Y-%m-%d %H:%M:%S")

    # region --------------- convert methods ---------------------------
    @staticmethod
    def date_from_number(value: numbers.Number) -> datetime.datetime:
        """
        Converts a float value to corresponding datetime instance.

        Args:
            value (Number): number to convert to date and time

        Raises:
            TypeError: If value is not correct type.

        Returns:
            datetime: Date time instance on success; Otherwise, None
        """
        if not isinstance(value, numbers.Real):
            raise TypeError(f"Incorrect type. Excpected 'Number' got {type(value).__name__}")
        delta = datetime.timedelta(days=value)
        dnull = cast(datetime.datetime, mLo.Lo.null_date)
        return dnull + delta

    @staticmethod
    def date_to_number(date: datetime.datetime | datetime.date) -> float:
        """
        Converts a date or datetime instance to a corresponding float value.

        Args:
            date (datetime | date): date or date time to convert to float

        Raises:
            TypeError: If date is not a correct type.

        Returns:
            float: date as float on success; Otherwise, None
        """
        dnull = cast(datetime.datetime, mLo.Lo.null_date)
        if isinstance(date, datetime.datetime):
            delta = date - dnull
        elif isinstance(date, datetime.date):
            delta = date - dnull.date()
        else:
            raise TypeError(f"Incorrect type. Excpected 'date' or 'datetime' got {type(date).__name__}")
        return delta.days + delta.seconds / (24.0 * 60 * 60)

    @staticmethod
    def time_from_number(value: numbers.Number) -> datetime.time | None:
        """
        Converts a float value to corresponding time instance.

        Only Hours, Minutes and seconds are are used in conversion.

        Args:
            value (Number): Number such a float or int to convert

        Raises:
            TypeError: If value is not correct type.

        Returns:
            time | None: Value as time on success; Otherwise, None
        """
        if not isinstance(value, numbers.Real):
            raise TypeError(f"Incorrect type. Excpected 'Number' got {type(value).__name__}")
        delta = datetime.timedelta(days=value)
        minutes, second = divmod(delta.seconds, 60)
        hour, minute = divmod(minutes, 60)
        return datetime.time(hour, minute, second, tzinfo=datetime.timezone.utc)

    @staticmethod
    def time_to_number(time: datetime.time) -> float:
        """
        Converts a time instance to a corresponding float value.

        Only Hours, Minutes and seconds are are used in conversion.

        Args:
            time (datetime.time): time to convert

        Raises:
            TypeError: If date is not a correct type.

        Returns:
            float: time as float on success; Otherwise, None
        """
        if not isinstance(time, datetime.time):
            raise TypeError(f"Incorrect type. Excpected 'Number' got {type(time).__name__}")
        return ((time.second / 60.0 + time.minute) / 60.0 + time.hour) / 24.0

    @staticmethod
    def date_time_str(dt: datetime.datetime) -> str:
        """
        returns a formatted date and time as string

        Args:
            dt (datetime): date time

        Returns:
            str: formatted date string such as ``Jun 05, 2022 20:15``
        """
        return dt.strftime("%b %d, %Y %H:%M")

    @staticmethod
    def uno_dt_to_dt(uno_dt: UnoDateTime) -> datetime.datetime:
        """
        Converts a uno DateTime struct to a datetime instance

        Args:
            uno_dt (UnoDateTime): uno Datetime struct

        Returns:
            datetime.datetime: Python DateTime
        """
        if uno_dt.Year <= 0 or uno_dt.Month <= 0 or uno_dt.Day <= 0:
            return mLo.Lo.null_date
        td = datetime.datetime(
            year=uno_dt.Year,
            month=uno_dt.Month,
            day=uno_dt.Day,
            hour=uno_dt.Hours,
            minute=uno_dt.Minutes,
            second=uno_dt.Seconds,
            microsecond=0 if uno_dt.NanoSeconds == 0 else int(uno_dt.NanoSeconds / 1000),
            tzinfo=datetime.timezone.utc if uno_dt.IsUTC else None,
        )
        return td

    @classmethod
    def str_date_time(cls, uno_dt: UnoDateTime) -> str:
        """
        returns a formatted date and time as string

        Args:
            uno_dt (datetime): date time

        Returns:
            str: formatted date string such as ``Jun 05, 2022 20:15``
            or empty string if ``uno_dt`` is null.
        """
        dt = cls.uno_dt_to_dt(uno_dt)
        if dt == mLo.Lo.null_date:
            return ""
        return cls.date_time_str(dt)

    # endregion ------------ convert methods ---------------------------
