from __future__ import annotations

from datetime import datetime, timedelta
import sys
import types

from dateutil.tz import tzlocal

# Stub win32com for environments where pywin32 is not installed (CI/tests).
if "win32com" not in sys.modules:
    win32com_module = types.ModuleType("win32com")
    win32com_client_module = types.ModuleType("win32com.client")
    win32com_client_module.Dispatch = lambda *args, **kwargs: None
    win32com_module.client = win32com_client_module
    sys.modules["win32com"] = win32com_module
    sys.modules["win32com.client"] = win32com_client_module

from src.outlook.client import OutlookClient
import src.outlook.client as outlook_client_module


class FakeMailItem:
    def __init__(self, subject: str, received: datetime, unread: bool) -> None:
        self.Subject = subject
        self.ReceivedTime = received
        self.UnRead = unread


class FakeItemsCollection:
    def __init__(self, items: list[FakeMailItem]) -> None:
        self._items = list(items)

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, i: int) -> FakeMailItem:
        return self._items[i - 1]

    def Sort(self, field: str, descending: bool) -> None:
        if field == "[ReceivedTime]":
            self._items.sort(key=lambda x: x.ReceivedTime, reverse=descending)

    def Restrict(self, filter_str: str) -> "FakeItemsCollection":
        filtered = self._items
        clauses = [x.strip() for x in filter_str.split("AND")]

        for clause in clauses:
            if clause.startswith("[ReceivedTime] >="):
                raw = clause.split(">=", 1)[1].strip().strip("'")
                since = datetime.strptime(raw, "%m/%d/%Y %I:%M %p")
                filtered = [
                    item
                    for item in filtered
                    if item.ReceivedTime.replace(tzinfo=None) >= since
                ]
            elif clause.startswith("[ReceivedTime] <"):
                raw = clause.split("<", 1)[1].strip().strip("'")
                until_exclusive = datetime.strptime(raw, "%m/%d/%Y %I:%M %p")
                filtered = [
                    item
                    for item in filtered
                    if item.ReceivedTime.replace(tzinfo=None) < until_exclusive
                ]
            elif clause == "[UnRead] = True":
                filtered = [item for item in filtered if item.UnRead]

        return FakeItemsCollection(filtered)


class FakeFolder:
    def __init__(self, items: list[FakeMailItem]) -> None:
        self.Items = FakeItemsCollection(items)


def _build_client_without_com() -> OutlookClient:
    # Avoid COM initialization for unit tests.
    return OutlookClient.__new__(OutlookClient)


def test_iter_items_filters_unread_only() -> None:
    now = datetime.now(tzlocal())
    folder = FakeFolder(
        [
            FakeMailItem("mail-1", now - timedelta(hours=1), unread=True),
            FakeMailItem("mail-2", now - timedelta(hours=2), unread=False),
            FakeMailItem("mail-3", now - timedelta(hours=3), unread=True),
        ]
    )
    client = _build_client_without_com()

    out = list(client.iter_items(folder, days_back=0, only_unread=True))

    assert [x.Subject for x in out] == ["mail-1", "mail-3"]


def test_iter_items_filters_days_back(monkeypatch) -> None:
    fixed_now = datetime(2026, 2, 25, 12, 0, tzinfo=tzlocal())

    class FixedDateTime(datetime):
        @classmethod
        def now(cls, tz=None):
            return fixed_now if tz is not None else fixed_now.replace(tzinfo=None)

    monkeypatch.setattr(outlook_client_module, "datetime", FixedDateTime)

    folder = FakeFolder(
        [
            FakeMailItem("today", fixed_now - timedelta(hours=2), unread=True),
            FakeMailItem("yesterday", fixed_now - timedelta(days=1), unread=True),
            FakeMailItem("old", fixed_now - timedelta(days=5), unread=True),
        ]
    )
    client = _build_client_without_com()

    out = list(client.iter_items(folder, days_back=2, only_unread=False))

    assert [x.Subject for x in out] == ["today", "yesterday"]


def test_print_today_and_yesterday_mails(monkeypatch) -> None:
    fixed_now = datetime(2026, 2, 25, 12, 0, tzinfo=tzlocal())

    class FixedDateTime(datetime):
        @classmethod
        def now(cls, tz=None):
            return fixed_now if tz is not None else fixed_now.replace(tzinfo=None)

    monkeypatch.setattr(outlook_client_module, "datetime", FixedDateTime)

    folder = FakeFolder(
        [
            FakeMailItem("hoy-1", fixed_now - timedelta(hours=1), unread=True),
            FakeMailItem("hoy-2", fixed_now - timedelta(hours=4), unread=False),
            FakeMailItem("ayer-1", fixed_now - timedelta(days=1, hours=2), unread=True),
            FakeMailItem("antiguo", fixed_now - timedelta(days=5), unread=True),
        ]
    )
    client = _build_client_without_com()

    all_recent = list(client.iter_items(folder, days_back=2, only_unread=False))
    today = [m.Subject for m in all_recent if m.ReceivedTime.date() == fixed_now.date()]
    yesterday = [m.Subject for m in all_recent if m.ReceivedTime.date() == (fixed_now - timedelta(days=1)).date()]

    print(f"[TEST] Correos de hoy ({fixed_now.date()}): {today}")
    print(f"[TEST] Correos de ayer ({(fixed_now - timedelta(days=1)).date()}): {yesterday}")

    assert today == ["hoy-1", "hoy-2"]
    assert yesterday == ["ayer-1"]


def test_iter_items_filters_exact_date_range(monkeypatch) -> None:
    fixed_now = datetime(2026, 2, 25, 12, 0, tzinfo=tzlocal())

    class FixedDateTime(datetime):
        @classmethod
        def now(cls, tz=None):
            return fixed_now if tz is not None else fixed_now.replace(tzinfo=None)

    monkeypatch.setattr(outlook_client_module, "datetime", FixedDateTime)

    folder = FakeFolder(
        [
            FakeMailItem("day-25", fixed_now - timedelta(hours=2), unread=True),
            FakeMailItem("day-24", fixed_now - timedelta(days=1, hours=1), unread=True),
            FakeMailItem("day-23", fixed_now - timedelta(days=2), unread=True),
        ]
    )
    client = _build_client_without_com()

    out = list(
        client.iter_items(
            folder,
            days_back=0,
            only_unread=False,
            date_from=datetime(2026, 2, 24).date(),
            date_to=datetime(2026, 2, 25).date(),
        )
    )

    assert [x.Subject for x in out] == ["day-25", "day-24"]
