# coding: utf-8
from __future__ import annotations
import datetime
import time
from typing import cast, Any, Tuple
import uno  # pylint: disable=unused-import
from ooo.dyn.util.date_time import DateTime as UnoDateTime
from ooo.dyn.util.date import Date as UnoDate
from ooo.dyn.util.time import Time as UnoTime
from ooodev.loader import lo as mLo


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
        dt = datetime.datetime.now(tz) if tz is not None else datetime.datetime.now()
        return dt.strftime("%Y-%m-%d %H:%M:%S")

    # region --------------- convert methods ---------------------------
    @staticmethod
    def date_from_number(value: int | float) -> datetime.datetime:
        """
        Converts a float value to corresponding datetime instance.

        Args:
            value (Number): number to convert to date and time

        Raises:
            TypeError: If value is not correct type.

        Returns:
            datetime: Date time instance on success; Otherwise, None
        """
        if not isinstance(value, (int, float)):
            raise TypeError(f"Incorrect type. Expected int or float got {type(value).__name__}")
        delta = datetime.timedelta(days=value)
        null_date = cast(datetime.datetime, mLo.Lo.null_date)
        return null_date + delta

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
        null_date = cast(datetime.datetime, mLo.Lo.null_date)
        if isinstance(date, datetime.datetime):
            delta = date - null_date
        elif isinstance(date, datetime.date):
            delta = date - null_date.date()  # pylint: disable=no-member
        else:
            raise TypeError(f"Incorrect type. Expected 'date' or 'datetime' got {type(date).__name__}")
        return delta.days + delta.seconds / (24.0 * 60 * 60)

    @staticmethod
    def time_from_number(value: int | float) -> datetime.time | None:
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
        if not isinstance(value, (int, float)):
            raise TypeError(f"Incorrect type. Expected int or float got {type(value).__name__}")
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
            raise TypeError(f"Incorrect type. Expected 'Number' got {type(time).__name__}")
        return ((time.second / 60.0 + time.minute) / 60.0 + time.hour) / 24.0

    @staticmethod
    def date_time_str(dt: datetime.datetime) -> str:
        """
        Returns a formatted date and time as string.

        |lo_safe|

        Args:
            dt (datetime): date time

        Returns:
            str: formatted date string such as ``Jun 05, 2022 20:15``
        """
        return dt.strftime("%b %d, %Y %H:%M")

    @staticmethod
    def uno_dt_to_dt(uno_dt: UnoDateTime) -> datetime.datetime:
        """
        Converts a uno DateTime struct to a datetime instance.

        |lo_safe|

        Args:
            uno_dt (UnoDateTime): uno Datetime struct

        Returns:
            datetime.datetime: Python DateTime
        """
        if uno_dt.Year <= 0 or uno_dt.Month <= 0 or uno_dt.Day <= 0:
            return mLo.Lo.null_date
        return datetime.datetime(
            year=uno_dt.Year,
            month=uno_dt.Month,
            day=uno_dt.Day,
            hour=uno_dt.Hours,
            minute=uno_dt.Minutes,
            second=uno_dt.Seconds,
            microsecond=0 if uno_dt.NanoSeconds == 0 else int(uno_dt.NanoSeconds / 1000),
            tzinfo=datetime.timezone.utc if uno_dt.IsUTC else None,
        )

    @staticmethod
    def uno_date_to_date(uno_date: UnoDate) -> datetime.datetime:
        """
        Converts a uno Date struct to a datetime instance

        Args:
            uno_date (UnoDate): uno Date struct

        Returns:
            datetime.datetime: Python DateTime
        """

        if uno_date.Year <= 0 or uno_date.Month <= 0 or uno_date.Day <= 0:
            return mLo.Lo.null_date
        return datetime.datetime(
            year=uno_date.Year, month=uno_date.Month, day=uno_date.Day, hour=0, minute=0, second=0, microsecond=0
        )

    @staticmethod
    def uno_time_to_date_time(uno_time: UnoTime) -> datetime.datetime:
        """
        Converts a uno Time struct to a datetime instance

        Args:
            uno_time (UnoTime): uno Time struct

        Returns:
            datetime.datetime: Python DateTime
        """
        # pylint: disable=no-member
        null_date = mLo.Lo.null_date
        dt = datetime.datetime(
            year=null_date.year,
            month=null_date.month,
            day=null_date.day,
            hour=uno_time.Hours,
            minute=uno_time.Minutes,
            second=uno_time.Seconds,
            microsecond=0 if uno_time.NanoSeconds == 0 else int(uno_time.NanoSeconds / 1000),
            tzinfo=datetime.timezone.utc if uno_time.IsUTC else None,
        )

        return dt

    @staticmethod
    def uno_time_to_time(uno_time: UnoTime) -> datetime.time:
        """
        Converts a uno Time struct to a time instance

        Args:
            uno_time (UnoTime): uno Time struct

        Returns:
            datetime.time: Python Time
        """
        tm = datetime.time(
            hour=uno_time.Hours,
            minute=uno_time.Minutes,
            second=uno_time.Seconds,
            microsecond=0 if uno_time.NanoSeconds == 0 else int(uno_time.NanoSeconds / 1000),
            tzinfo=datetime.timezone.utc if uno_time.IsUTC else None,
        )

        return tm

    @classmethod
    def time_to_uno_time(cls, time: datetime.time) -> UnoTime:
        """
        Converts a python time to  UNO Time struct instance

        Args:
            time (UnoTime): Python time

        Returns:
            Time: UNO Time struct
        """
        dt = cls.date_to_uno_date_time(time)
        return UnoTime(
            dt.NanoSeconds,
            dt.Seconds,
            dt.Minutes,
            dt.Hours,
            dt.IsUTC,
        )

    @staticmethod
    def date_to_uno_date_time(date: Any) -> UnoDateTime:
        """
        Converts a date representation into the com.sun.star.util.DateTime date format
        Acceptable boundaries: ``year >= 1900`` and ``<= 32767``

        Args:
            date (Any): ``datetime.datetime``, ``datetime.date``, ``datetime.time``, ``float`` (time.time) or ``time.struct_time``

        Returns:
            DateTime: A ``com.sun.star.util.DateTime`` if conversion was successful; Otherwise, ``date``
        """
        uno_date = UnoDateTime()
        uno_date.Year = 1899
        uno_date.Month = 12
        uno_date.Day = 30
        uno_date.Hours = 0
        uno_date.Minutes = 0
        uno_date.Seconds = 0
        uno_date.NanoSeconds = 0
        uno_date.IsUTC = False

        if isinstance(date, float):
            date = time.localtime(date)
        if isinstance(date, time.struct_time):
            if 1900 <= date[0] <= 32767:
                (
                    uno_date.Year,
                    uno_date.Month,
                    uno_date.Day,
                    uno_date.Hours,
                    uno_date.Minutes,
                    uno_date.Seconds,
                ) = date[0:6]
            else:  # Copy only the time related part
                uno_date.Hours, uno_date.Minutes, uno_date.Seconds = cast(Tuple[int, int, int], date[3:3])
        elif isinstance(date, (datetime.datetime, datetime.date, datetime.time)):
            if isinstance(date, (datetime.datetime, datetime.date)):
                if 1900 <= date.year <= 32767:
                    uno_date.Year, uno_date.Month, uno_date.Day = (
                        date.year,
                        date.month,
                        date.day,
                    )
            if isinstance(date, (datetime.datetime, datetime.time)):
                (
                    uno_date.Hours,
                    uno_date.Minutes,
                    uno_date.Seconds,
                    uno_date.NanoSeconds,
                ) = (date.hour, date.minute, date.second, date.microsecond * 1000)
        else:
            return date  # Not recognized as a date
        return uno_date

    @classmethod
    def date_to_uno_date(cls, date: Any) -> UnoDate:
        """
        Converts a date representation into the com.sun.star.util.DateTime date format
        Acceptable boundaries: ``year >= 1900`` and ``<= 32767``

        Args:
            date (Any): ``datetime.datetime``, ``datetime.date``, ``datetime.time``, ``float`` (time.time) or ``time.struct_time``

        Returns:
            Date: A ``com.sun.star.util.Time`` if conversion was successful; Otherwise, ``date``
        """
        try:
            dt = cls.date_to_uno_date_time(date)
            return UnoDate(dt.Day, dt.Month, dt.Year)
        except Exception:
            return date

    @classmethod
    def str_date_time(cls, uno_dt: UnoDateTime) -> str:
        """
        Returns a formatted date and time as string.

        |lo_safe|

        Args:
            uno_dt (datetime): date time

        Returns:
            str: formatted date string such as ``Jun 05, 2022 20:15``
            or empty string if ``uno_dt`` is null.
        """
        dt = cls.uno_dt_to_dt(uno_dt)
        return "" if dt == mLo.Lo.null_date else cls.date_time_str(dt)

    # endregion ------------ convert methods ---------------------------
