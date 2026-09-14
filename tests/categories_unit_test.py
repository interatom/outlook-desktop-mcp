"""
Outlook Desktop MCP - Category Unit Tests
===========================================
Pure-logic tests for the color-category plumbing: the split/match helpers,
the replace/add/remove modes of set_category, and the presence of the
`categories` field in the event summary. The COM bridge is stubbed, so this
needs NO COM and no running Outlook.

Run: .venv\\Scripts\\python tests\\categories_unit_test.py
"""
import asyncio
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp.server import (  # noqa: E402
    _category_separator,
    _join_categories,
    _matches_any_category,
    _split_categories,
    mcp,
)
from outlook_desktop_mcp.utils.formatting import (  # noqa: E402
    format_event_full,
    format_event_summary,
)
import outlook_desktop_mcp.server as server  # noqa: E402


def log(msg):
    print(msg, file=sys.stderr, flush=True)


class FakeItem:
    """Minimal stand-in for an Outlook item that carries categories."""

    def __init__(self, categories="", subject="Testtermin"):
        self.Categories = categories
        self.Subject = subject
        self.saved = 0

    def Save(self):
        self.saved += 1


class FakeNamespace:
    def __init__(self, item):
        self._item = item

    def GetItemFromID(self, entry_id, store_id=None):
        return self._item


def call_tool(name, args, item):
    """Invoke a registered MCP tool with the COM bridge stubbed out.

    Runs the tool's inner closure against FakeNamespace instead of a real
    Outlook, so the production code path -- not a copy of it -- is measured.
    """
    original = server.bridge.call

    async def fake_call(func, *a, **kw):
        return func(None, FakeNamespace(item), *a, **kw)

    server.bridge.call = fake_call
    try:
        return asyncio.run(mcp.call_tool(name, args))
    finally:
        server.bridge.call = original


# --- helpers ---

def test_split_trims_and_drops_empties():
    assert _split_categories(" Kunde A , , Reise ") == ["Kunde A", "Reise"]
    assert _split_categories("") == []
    assert _split_categories(None) == []
    log("  ' Kunde A , , Reise ' -> ['Kunde A', 'Reise']")


def test_split_accepts_the_semicolon_outlook_actually_writes():
    """Regression guard for the locale separator.

    Outlook joins Categories with the Windows LIST SEPARATOR, not a comma:
    measured 2026-09-14, an event created with "Internal, MCP-Smoketest" read
    back as "Internal; MCP-Smoketest". Splitting on "," alone produced ONE
    token, which made every add/remove a silent no-op and every category
    filter return []. Both were observed live before the fix.
    """
    assert _split_categories("Internal; MCP-Smoketest") == [
        "Internal",
        "MCP-Smoketest",
    ]
    assert _split_categories("A; B, C") == ["A", "B", "C"]  # mixed, both accepted
    log("  'Internal; MCP-Smoketest' -> two names, not one")


def test_separator_is_not_hardcoded():
    """The emitted separator must come from the system, not from a literal."""
    sep = _category_separator()
    assert sep.rstrip() in (",", ";"), sep
    assert sep.endswith(" "), repr(sep)
    round_tripped = _split_categories(_join_categories(["A", "B"]))
    assert round_tripped == ["A", "B"], round_tripped
    log(f"  separator on this system: {sep!r}, round-trip clean")


def test_match_survives_the_semicolon_form():
    """The filter has to work on what Outlook returns, not on what we wrote."""
    assert _matches_any_category("Internal; MCP-Smoketest", ["internal"]) is True
    assert _matches_any_category("Internal; MCP-Smoketest", ["Urlaub"]) is False
    log("  filter matches against the semicolon form")


def test_match_is_case_insensitive_and_any_of():
    assert _matches_any_category("Kunde A, Reise", ["reise"]) is True
    assert _matches_any_category("Kunde A, Reise", ["Urlaub", "KUNDE A"]) is True
    assert _matches_any_category("Kunde A", ["Reise"]) is False
    assert _matches_any_category("", ["Reise"]) is False
    log("  any-of, case-insensitive")


def test_empty_filter_matches_everything():
    """No category argument must not silently hide uncategorized events."""
    assert _matches_any_category("", []) is True
    assert _matches_any_category("Kunde A", []) is True
    log("  empty filter is a pass-through, not a 'has categories' filter")


def test_match_does_not_do_substrings():
    """'Kunde A' must not match a request for 'Kunde'."""
    assert _matches_any_category("Kunde A", ["Kunde"]) is False
    log("  'Kunde' does not match the item's 'Kunde A'")


# --- set_category modes ---

def test_replace_overwrites():
    item = FakeItem("Alt")
    call_tool("set_category", {"entry_id": "X", "categories": "Neu"}, item)
    assert item.Categories == "Neu", item.Categories
    assert item.saved == 1
    log(f"  'Alt' -> '{item.Categories}' (default mode)")


def test_replace_with_empty_string_clears():
    item = FakeItem("Kunde A, Reise")
    call_tool("set_category", {"entry_id": "X", "categories": ""}, item)
    assert item.Categories == "", item.Categories
    log("  empty string clears all categories")


def test_add_keeps_existing():
    item = FakeItem("Kunde A")
    call_tool(
        "set_category",
        {"entry_id": "X", "categories": "Reise", "mode": "add"},
        item,
    )
    assert item.Categories == _join_categories(["Kunde A", "Reise"]), item.Categories
    log(f"  add -> '{item.Categories}'")


def test_add_is_idempotent_and_keeps_existing_spelling():
    """Re-adding must not duplicate, and must not re-case what is there."""
    item = FakeItem("Kunde A")
    call_tool(
        "set_category",
        {"entry_id": "X", "categories": "kunde a", "mode": "add"},
        item,
    )
    assert item.Categories == "Kunde A", item.Categories
    log("  adding 'kunde a' to 'Kunde A' is a no-op")


def test_remove_subtracts_only_named():
    item = FakeItem("Kunde A, Reise, Privat")
    call_tool(
        "set_category",
        {"entry_id": "X", "categories": "reise", "mode": "remove"},
        item,
    )
    assert item.Categories == _join_categories(["Kunde A", "Privat"]), item.Categories
    log(f"  remove -> '{item.Categories}'")


def test_remove_of_absent_category_is_a_no_op():
    item = FakeItem("Kunde A")
    call_tool(
        "set_category",
        {"entry_id": "X", "categories": "Urlaub", "mode": "remove"},
        item,
    )
    assert item.Categories == "Kunde A", item.Categories
    log("  removing a category the item does not carry changes nothing")


def test_written_value_uses_the_system_separator():
    """A round-trip must not reformat the field -- on ANY locale."""
    item = FakeItem("")
    call_tool(
        "set_category",
        {"entry_id": "X", "categories": "A,B,C", "mode": "replace"},
        item,
    )
    assert item.Categories == _join_categories(["A", "B", "C"]), item.Categories
    assert _split_categories(item.Categories) == ["A", "B", "C"]
    log(f"  'A,B,C' -> '{item.Categories}'")


def test_add_and_remove_work_on_the_semicolon_form():
    """The live failure, as a test: both were silent no-ops before the fix."""
    item = FakeItem("Internal; MCP-Smoketest")
    call_tool(
        "set_category",
        {"entry_id": "X", "categories": "Internal", "mode": "remove"},
        item,
    )
    assert _split_categories(item.Categories) == ["MCP-Smoketest"], item.Categories

    call_tool(
        "set_category",
        {"entry_id": "X", "categories": "Reise", "mode": "add"},
        item,
    )
    assert _split_categories(item.Categories) == ["MCP-Smoketest", "Reise"], item.Categories
    log(f"  remove then add -> '{item.Categories}'")


def test_invalid_mode_fails_loudly():
    """A typo must not silently fall through to replace and wipe the item.

    Like every tool in this server, set_category reports failure as a text
    result rather than raising, so assert on the message AND on the item
    still carrying what it had.
    """
    item = FakeItem("Kunde A")
    result = str(
        call_tool(
            "set_category",
            {"entry_id": "X", "categories": "Reise", "mode": "append"},
            item,
        )
    )
    assert "append" in result, result
    assert "replace, add, remove" in result, result
    assert item.Categories == "Kunde A", "item must be untouched"
    assert item.saved == 0, "invalid mode must not Save()"
    log("  rejected, item untouched, no Save()")


# --- summary field ---

class FakeAppointment:
    EntryID = "0000ABCD"
    Subject = "Kickoff"
    Start = "2026-09-14 10:00"
    End = "2026-09-14 11:00"
    Duration = 60
    Location = "Raum 1"
    Organizer = "Someone"
    IsRecurring = False
    AllDayEvent = False
    BusyStatus = 2
    MeetingStatus = 0
    RequiredAttendees = ""
    OptionalAttendees = ""
    Categories = "Kunde A, Reise"
    Body = "agenda"
    ReminderSet = False
    ResponseStatus = 0


def test_summary_carries_categories():
    """list_events/search_events must not need an extra get_event per item."""
    out = format_event_summary(FakeAppointment())
    assert out["categories"] == "Kunde A, Reise", out
    log(f"  summary categories: '{out['categories']}'")


def test_full_does_not_duplicate_categories():
    out = format_event_full(FakeAppointment())
    assert list(out.keys()).count("categories") == 1, out.keys()
    assert out["categories"] == "Kunde A, Reise", out
    log("  full view inherits the field instead of setting it twice")


def main():
    tests = [
        ("Split trims and drops empties", test_split_trims_and_drops_empties),
        ("Match is case-insensitive any-of", test_match_is_case_insensitive_and_any_of),
        ("Empty filter matches everything", test_empty_filter_matches_everything),
        ("Match does not do substrings", test_match_does_not_do_substrings),
        ("Replace overwrites", test_replace_overwrites),
        ("Empty string clears", test_replace_with_empty_string_clears),
        ("Add keeps existing", test_add_keeps_existing),
        ("Add is idempotent", test_add_is_idempotent_and_keeps_existing_spelling),
        ("Remove subtracts only named", test_remove_subtracts_only_named),
        ("Remove of absent is no-op", test_remove_of_absent_category_is_a_no_op),
        ("Written value uses system separator", test_written_value_uses_the_system_separator),
        ("Add/remove on semicolon form", test_add_and_remove_work_on_the_semicolon_form),
        ("Split accepts semicolon", test_split_accepts_the_semicolon_outlook_actually_writes),
        ("Separator is not hardcoded", test_separator_is_not_hardcoded),
        ("Match survives semicolon form", test_match_survives_the_semicolon_form),
        ("Invalid mode fails loudly", test_invalid_mode_fails_loudly),
        ("Summary carries categories", test_summary_carries_categories),
        ("Full does not duplicate", test_full_does_not_duplicate_categories),
    ]
    failed = 0
    for name, fn in tests:
        try:
            log(f"[ RUN  ] {name}")
            fn()
            log(f"[  OK  ] {name}")
        except Exception as e:
            failed += 1
            log(f"[ FAIL ] {name}: {type(e).__name__}: {e}")
    log("")
    log(f"{len(tests) - failed}/{len(tests)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
