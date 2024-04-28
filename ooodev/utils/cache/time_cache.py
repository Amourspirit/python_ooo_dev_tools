from __future__ import annotations
import threading
from typing import Any
from datetime import datetime, timedelta, timezone
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.events.partial.events_partial import EventsPartial


class TimeCache(EventsPartial):
    """
    Time based Cache.

    Cached items expire after a specified time. If ``cleanup_interval`` is set, then the cache is cleaned up at regular intervals;
    Otherwise, the cache is only cleaned up when an item is accessed.

    Each time an element is accessed, the timestamp is updated. If the element has expired, it is removed from the cache.

    When an item expires, the event ``cache_items_expired`` is triggered.
    This event is called on a separate thread. for this reason it is important to make sure that the event handler is thread safe.

    Example:

        .. code-block:: python

            import threading
            from ooodev.utils.cache.time_cache import TimeCache

            LOCK = threading.Lock()

            def on_items_expired(source, event):
                with LOCK:
                    keys = event.event_data.keys
                    for key in keys:
                        print(f"Expired: {key}")

            cache = TimeCache(60.0)  # 60 seconds
            cache.subscribe_event("cache_items_expired", on_items_expired)
            cache["key"] = "value"
            value = cache["key"]
    """

    def __init__(self, seconds: float, cleanup_interval: float = 60.0) -> None:
        """
        Time based Cache.

        Args:
            seconds (float): Cache expiration time in seconds.
            cleanup_interval (float, optional): Cache cleanup interval in seconds.
                If set to ``0`` then the cleanup is disabled. Defaults to ``60.0``.
        """
        EventsPartial.__init__(self)
        self._lock = threading.Lock()
        self._delta = timedelta(seconds=max(seconds, 0))
        self._seconds = self._delta.total_seconds()
        self._timeout = max(cleanup_interval, 0)
        self._expiration_time = datetime.now(timezone.utc) + self._delta
        self._cache = {}
        self._timer = None
        self._fn_clear_expired = self.clear_expired
        self._fn_trigger_event = self.trigger_event
        self.start_timer()

    def clear(self) -> None:
        """
        Clear cache.
        """
        self.stop_timer()
        self._cache.clear()
        self.stop_timer()

    def get(self, key: Any) -> Any:
        """
        Get value by key.

        Args:
            key (Any): Any Hashable object.

        Returns:
            Any: Value or ``None`` if not found.

        Note:
            The ``get`` method is an alias for the ``__getitem__`` method.
            So you can use ``cache_inst.get(key)`` or ``cache_inst[key]`` interchangeably.
        """
        return self[key]

    def put(self, key: Any, value: Any) -> None:
        """
        Put value by key.

        Args:
            key (Any): Any Hashable object.
            value (Any): Any object.

        Note:
            The ``put`` method is an alias for the ``__setitem__`` method.
            So you can use ``cache_inst.put(key, value)`` or ``cache_inst[key] = value`` interchangeably.
        """
        self[key] = value

    def remove(self, key: Any) -> None:
        """
        Remove key.

        Args:
            key (Any): Any Hashable object.

        Note:
            The ``remove`` method is an alias for the ``__delitem__`` method.
            So you can use ``cache_inst.remove(key)`` or ``del cache_inst[key]`` interchangeably.
        """
        del self[key]

    def start_timer(self) -> bool:
        """
        Start the timer.

        This only applies if the cleanup interval is set.

        Returns:
            bool: ``True`` if the timer is started, ``False`` if the timer is already running.

        Note:
            Triggers the event ``time_cache_timer_started``.
        """
        if self.seconds > 0 and self._timeout > 0 and self._timer is None:
            self._timer = threading.Timer(self._timeout, self._fn_clear_expired)
            self._timer.daemon = True  # important for the timer to stop when the main thread exits
            self._timer.start()
            eargs = EventArgs(self)
            eargs.event_data = DotDict(timer=self._timer)
            self.trigger_event("time_cache_timer_started", eargs)
            return True
        return False

    def stop_timer(self) -> bool:
        """
        Stop the timer.

        This only applies if the cleanup interval is set.

        Returns:
            bool: ``True`` if the timer is stopped, ``False`` if the timer was not running.

        Note:
            Triggers the event ``time_cache_timer_stopped``.
        """
        if self._timer is not None:
            self._timer.cancel()
            self._timer = None
            eargs = EventArgs(self)
            self.trigger_event("time_cache_timer_stopped", eargs)
            return True
        return False

    def clear_expired(self) -> None:
        """
        Clear expired items from the cache.

        Note:
            Triggers the event ``cache_items_expired``, on a new thread.

            The event args ``event_data`` is a ``DotDict`` instance that contains the ``keys`` as a list of the items that were removed.
        """

        self.stop_timer()
        with self._lock:
            if not self._cache:
                self.start_timer()
                return

            seconds = self.seconds
            del_keys = [
                key
                for key, (_, timestamp) in self._cache.items()
                if (datetime.now(timezone.utc) - timestamp).total_seconds() >= seconds
            ]
            for key in del_keys:
                del self._cache[key]
        if del_keys:
            eargs = EventArgs(self)
            eargs.event_data = DotDict(keys=del_keys)

            thread = threading.Thread(target=self._fn_trigger_event, args=("cache_items_expired", eargs), daemon=True)
            thread.start()
            thread.join()  # wait for the thread to finish before starting timer again.
        self.start_timer()

    # region Dunder Methods
    def __getitem__(self, key: Any) -> Any:
        if key is None:
            raise TypeError("Key must not be None.")
        if key in self._cache:
            value, timestamp = self._cache[key]
            now_dt = datetime.now(timezone.utc)
            if (now_dt - timestamp).total_seconds() < self.seconds:
                # update timestamp
                self._cache[key] = (value, now_dt)
                return value
            else:
                del self._cache[key]  # remove expired item
        return None

    def __setitem__(self, key: Any, value: Any) -> None:
        """
        Set value by key.

        Args:
            key (Any): Any Hashable object.
            value (Any): Any object.

        Raises:
            TypeError: If key or value is ``None``.

        Note:
            Triggers the event ``cache_item_adding`` before adding the item.
            The Event is a ``CancelEventArgs`` and can be canceled.

            Triggers the event ``cache_item_added`` after adding the item.
            The Event is a ``EventArgs``.

            Triggers the event ``cache_item_updating`` before updating the item.
            The Event is a ``CancelEventArgs`` and can be canceled.

            Triggers the event ``cache_item_updated`` after updating the item.
            The Event is a ``EventArgs``.

            The event args ``event_data`` is a ``DotDict`` instance that contains the ``key``, ``value`` and ``is_new`` of the item being added or updated.
        """
        if key is None or value is None:
            raise TypeError("Key and value must not be None.")
        is_new = self[key] is None
        if is_new:
            cargs = CancelEventArgs(self)
            cargs.event_data = DotDict(key=key, value=value, is_new=is_new)
            self.trigger_event("cache_item_adding", cargs)
            if cargs.cancel:
                return
            eargs = EventArgs.from_args(cargs)
        else:
            cargs = CancelEventArgs(self)
            cargs.event_data = DotDict(key=key, value=value, is_new=is_new)
            self.trigger_event("cache_item_updating", cargs)
            if cargs.cancel:
                return
            eargs = EventArgs.from_args(cargs)
        self._cache[key] = (value, datetime.now(timezone.utc))

        if is_new:
            self.trigger_event("cache_item_added", eargs)
        else:
            self.trigger_event("cache_item_updated", eargs)

    def __contains__(self, key: Any) -> bool:
        return False if key is None else self[key] is not None

    def __delitem__(self, key: Any) -> None:
        """
        Remove key.

        Args:
            key (Any): Any Hashable object.

        Raises:
            TypeError: If key is ``None``.

        Note:
            Triggers the event ``cache_item_removing`` before removing the item.
            The Event is a ``CancelEventArgs`` and can be canceled.

            Triggers the event ``cache_item_removed`` after removing the item.
            The Event is a ``EventArgs``.

            The event args ``event_data`` is a ``DotDict`` instance that contains the key of the item being removed.
        """
        if key is None:
            raise TypeError("Key must not be None.")
        if key in self:
            cargs = CancelEventArgs(self)
            cargs.event_data = DotDict(key=key)
            self.trigger_event("cache_item_removing", cargs)
            if cargs.cancel:
                return
            del self._cache[key]
            eargs = EventArgs.from_args(cargs)
            self.trigger_event("cache_item_removed", eargs)

    def __repr__(self) -> str:
        return f"TimeBasedCache({self.seconds})"

    def __str__(self) -> str:
        return f"TimeBasedCache({self.seconds})"

    def __len__(self) -> int:
        return len(self._cache)

    def __del__(self):
        # not reliable but better than nothing
        self.stop_timer()

    # endregion Dunder Methods

    # region Properties
    @property
    def seconds(self) -> float:
        """
        Gets/Sets Cache expiration time in seconds.
        """
        return self._seconds

    @seconds.setter
    def seconds(self, value: float) -> None:
        self._delta = timedelta(seconds=value)
        self._seconds = self._delta.total_seconds()
        self._expiration_time = datetime.now(timezone.utc) + self._delta
        self.clear_expired()

    @property
    def cleanup_interval(self) -> float:
        """
        Gets/Sets Cache cleanup interval in seconds.
        """
        return self._timeout

    @cleanup_interval.setter
    def cleanup_interval(self, value: float) -> None:
        self._timeout = max(value, 0)
        if self._timeout > 0:
            self.start_timer()
        else:
            self.stop_timer()

    # endregion Properties
