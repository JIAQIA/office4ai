# filename: dispose.py
# @Time    : 2024/4/18 19:30
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
# Python equivalent of your JavaScript code
"""
Enables logging of potentially leaked disposables.

A disposable is considered leaked if it is not disposed or not registered as the child of another disposable.
This tracking is very simple an only works for classes that either extend Disposable or use a DisposableStore.
This means there are a lot of false positives.
"""

from collections.abc import Callable, Iterable
from typing import (
    Any,
    Optional,
    Protocol,
    runtime_checkable,
)

from typing_extensions import Self

TRACK_DISPOSABLES = False
disposable_tracker: Optional["DisposableTracker"] = None


@runtime_checkable
class DisposableProtocol(Protocol):
    def dispose(self) -> None: ...


class DisposableTracker:
    __is_disposable_tracked__: str = "__is_disposable_tracked__"

    def track_disposable(self, x: DisposableProtocol) -> None:
        import traceback

        stack = traceback.format_stack()
        if not getattr(x, self.__is_disposable_tracked__, None):
            print(stack)
            setattr(x, self.__is_disposable_tracked__, True)

    def set_parent(self, child: DisposableProtocol, parent: Optional[DisposableProtocol]) -> None:  # noqa
        if child and child is not Disposable.none:
            setattr(child, self.__is_disposable_tracked__, True)

    def mark_as_disposed(self, disposable: DisposableProtocol) -> None:
        if disposable and disposable is not Disposable.none:
            setattr(disposable, self.__is_disposable_tracked__, True)

    def mark_as_singleton(self, disposable: DisposableProtocol) -> None:
        pass


def set_disposable_tracker(tracker: DisposableTracker | None) -> None:
    """
    Args:
        tracker: An instance of the DisposableTracker class that will be assigned to the global variable
            disposable_tracker.

    Returns:
        None.
    """
    global disposable_tracker
    disposable_tracker = tracker


def track_disposable(x: DisposableProtocol) -> DisposableProtocol:
    global disposable_tracker
    if disposable_tracker:
        disposable_tracker.track_disposable(x)
    return x


if TRACK_DISPOSABLES:
    set_disposable_tracker(DisposableTracker())  # pragma: no cover


class DisposableStore(DisposableProtocol):
    DISABLE_DISPOSED_WARNING = False

    def __init__(self) -> None:
        self._to_dispose: set = set()
        self._is_disposed = False
        track_disposable(self)

    def dispose(self) -> None:
        if self._is_disposed:
            return
        mark_as_disposed(self)
        self._is_disposed = True
        self.clear()

    @property
    def is_disposed(self) -> bool:
        return self._is_disposed

    def clear(self) -> None:
        if not self._to_dispose:
            return
        try:
            dispose(self._to_dispose)
        finally:
            self._to_dispose.clear()

    def add(self, o: DisposableProtocol) -> DisposableProtocol:
        if not isinstance(o, DisposableProtocol):
            return o
        if o == self:
            raise Exception("Cannot register a disposable on itself!")
        set_parent_of_disposable(o, self)
        if self._is_disposed and not self.DISABLE_DISPOSED_WARNING:
            print(
                "Trying to add a disposable to a DisposableStore that has already been disposed of. "
                "The added object will be leaked!",
            )
        else:
            self._to_dispose.add(o)
        return o

    def delete_and_leak(self, o: DisposableProtocol) -> None:
        if o in self._to_dispose:
            self._to_dispose.remove(o)
            set_parent_of_disposable(o, None)


class Disposable(DisposableProtocol):
    none = None

    def __init__(self) -> None:
        self._store = DisposableStore()
        track_disposable(self)
        set_parent_of_disposable(self._store, self)

    def dispose(self) -> None:
        mark_as_disposed(self)
        self._store.dispose()

    def register(self, o: DisposableProtocol) -> DisposableProtocol:
        if o == self:
            raise Exception("Cannot register a disposable on itself!")
        return self._store.add(o)


class MutableDisposable:
    def __init__(self) -> None:
        self._is_disposed = False
        self._value: DisposableProtocol | None = None
        track_disposable(self)

    @property
    def value(self) -> DisposableProtocol | None:
        return None if self._is_disposed else self._value

    @value.setter
    def value(self, value: DisposableProtocol | None) -> None:
        if not self._is_disposed and value != self._value:
            if self._value:
                self._value.dispose()
            if value:
                set_parent_of_disposable(value, self)
            self._value = value

    def clear(self) -> None:
        self.value = None

    def dispose(self) -> None:
        self._is_disposed = True
        mark_as_disposed(self)
        if self._value:
            self._value.dispose()
        self._value = None


class RefCountedDisposable:
    def __init__(self, disposable: DisposableProtocol) -> None:
        self._disposable: DisposableProtocol = disposable
        self._counter = 1

    def acquire(self) -> Self:
        self._counter += 1
        return self

    def release(self) -> Self:
        self._counter -= 1
        if self._counter == 0:
            self._disposable.dispose()
        return self


class ImmortalReference:
    def __init__(self, obj: DisposableProtocol) -> None:
        self.object = obj  # pragma: no cover

    @staticmethod
    def dispose() -> None:
        pass


class DisposableMap(DisposableProtocol):
    def __init__(self) -> None:
        self._store: dict[str, DisposableProtocol] = {}
        self._is_disposed = False
        track_disposable(self)

    def dispose(self) -> None:
        mark_as_disposed(self)
        self._is_disposed = True
        self.clear_and_dispose_all()

    def clear_and_dispose_all(self) -> None:
        if not self._store:
            return  # pragma: no cover
        try:
            dispose(self._store.values())
        finally:
            self._store.clear()

    def get(self, key: str) -> DisposableProtocol | None:
        return self._store.get(key)

    def set(
        self,
        key: str,
        value: DisposableProtocol,
        skip_dispose_on_overwrite: bool = False,
    ) -> None:
        if self._is_disposed:
            print(
                "Trying to add a disposable to a DisposableMap that has already been disposed of. "
                "The added object will be leaked!",
            )
        if not skip_dispose_on_overwrite and key in self._store:
            self._store[key].dispose()
        self._store[key] = value

    def delete_and_dispose(self, key: str) -> None:
        value = self._store.get(key)
        if value:
            value.dispose()
        del self._store[key]

    def __iter__(self) -> Iterable[str]:
        return iter(self._store)  # pragma: no cover


def mark_as_disposed(disposable: DisposableProtocol) -> None:
    global disposable_tracker
    if disposable_tracker:
        disposable_tracker.mark_as_disposed(disposable)


def set_parent_of_disposable(child: DisposableProtocol, parent: DisposableProtocol | None) -> None:
    global disposable_tracker
    if disposable_tracker:
        disposable_tracker.set_parent(child, parent)


def set_parent_of_disposables(children: Iterable[DisposableProtocol], parent: DisposableProtocol) -> None:
    global disposable_tracker
    if disposable_tracker:
        for child in children:
            disposable_tracker.set_parent(child, parent)


def mark_as_singleton(singleton: "Disposable") -> "Disposable":
    global disposable_tracker
    if disposable_tracker:
        disposable_tracker.mark_as_singleton(singleton)
    return singleton


def is_disposable(thing: Any) -> bool:
    return isinstance(thing, DisposableProtocol)  # pragma: no cover


def dispose(
    arg: DisposableProtocol | Iterable[DisposableProtocol],
) -> DisposableProtocol | Iterable[DisposableProtocol]:
    if isinstance(arg, Iterable):
        errors = []
        for d in arg:
            if d:
                try:
                    d.dispose()
                except Exception as e:
                    errors.append(e)
        if len(errors) == 1:
            raise errors[0]
        elif len(errors) > 1:
            raise Exception("Encountered errors while disposing of store")
        return [] if isinstance(arg, list) else arg
    else:
        arg.dispose()
        return arg


def combined_disposable(*disposables: DisposableProtocol) -> DisposableProtocol:
    """
    Combine multiple disposable values into a single {@link IDisposable}.

    Args:
        *disposables:

    Returns:

    """
    parent = to_disposable(lambda: dispose(disposables))
    set_parent_of_disposables(disposables, parent)
    return parent


def to_disposable(fn: Callable) -> DisposableProtocol:
    class Self(Disposable):
        def dispose(self) -> None:
            mark_as_disposed(self)
            fn()

    return track_disposable(Self())
