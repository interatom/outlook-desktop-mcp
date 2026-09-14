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


class FakeCategory:
    def __init__(self, name, color=0):
        self.Name = name
        self.Color = color


class FakeCategories:
    """Stand-in for namespace.Categories, 1-BASED like the COM collection.

    Mirrors the two behaviours measured against Outlook on 2026-09-14:
    Add() raises on a duplicate name, Remove() accepts the name.
    """

    def __init__(self, names=()):
        self._items = [FakeCategory(n, i + 1) for i, n in enumerate(names)]

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]

    def Add(self, name, color=None):
        if any(c.Name.lower() == name.lower() for c in self._items):
            raise ValueError("Value does not fall within the expected range.")
        self._items.append(FakeCategory(name, 7 if color is None else color))

    def Remove(self, index):
        if isinstance(index, str):
            for i, c in enumerate(self._items):
                if c.Name == index:
                    del self._items[i]
                    return
            raise ValueError("no such category")
        del self._items[index - 1]

    def names(self):
        return [c.Name for c in self._items]


class FakeCatNamespace:
    def __init__(self, categories):
        self.Categories = categories


def _run_with_namespace(name, args, namespace):
    original = server.bridge.call

    async def fake_call(func, *a, **kw):
        return func(None, namespace, *a, **kw)

    server.bridge.call = fake_call
    try:
        return str(asyncio.run(mcp.call_tool(name, args)))
    finally:
        server.bridge.call = original


def call_tool(name, args, item):
    """Invoke a registered MCP tool with the COM bridge stubbed out.

    Runs the tool's inner closure against FakeNamespace instead of a real
    Outlook, so the production code path -- not a copy of it -- is measured.
    """
    return _run_with_namespace(name, args, FakeNamespace(item))


def call_cat_tool(name, args, cats):
    """Same, for the master-list tools, which talk to namespace.Categories."""
    return _run_with_namespace(name, args, FakeCatNamespace(cats))


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


# --- master list ---

def test_create_category_adds_with_named_color():
    cats = FakeCategories(["Internal"])
    out = call_cat_tool("create_category", {"name": "Reise", "color": "blue"}, cats)
    assert cats.names() == ["Internal", "Reise"], cats.names()
    assert cats.Item(2).Color == 8, cats.Item(2).Color  # olCategoryColorBlue
    assert "blue" in out, out
    log(f"  added 'Reise' as color 8; list now {cats.names()}")


def test_create_category_is_idempotent_not_an_error():
    """Add() raises on a duplicate -- pre-check, and answer without a COM error."""
    cats = FakeCategories(["Internal"])
    out = call_cat_tool("create_category", {"name": "internal"}, cats)
    assert cats.names() == ["Internal"], cats.names()
    assert "already exists" in out, out
    assert "unexpected error" not in out.lower(), out
    log("  duplicate (different case) reported, not raised, not duplicated")


def test_create_category_rejects_an_unknown_color():
    cats = FakeCategories([])
    out = call_cat_tool("create_category", {"name": "X", "color": "chartreuse"}, cats)
    assert cats.names() == [], cats.names()
    assert "chartreuse" in out and "dark_maroon" in out, out
    log("  invalid color names the valid set and adds nothing")


def test_rename_category_renames_only_the_list_entry():
    """Measured: items keep the old text, so the answer has to say so."""
    cats = FakeCategories(["Internal", "Reise"])
    out = call_cat_tool(
        "rename_category", {"name": "internal", "new_name": "Intern"}, cats
    )
    assert cats.names() == ["Intern", "Reise"], cats.names()
    assert "NOT retagged" in out, out
    log(f"  renamed in place -> {cats.names()}, warning present")


def test_rename_category_refuses_a_collision():
    cats = FakeCategories(["Internal", "Reise"])
    out = call_cat_tool(
        "rename_category", {"name": "Internal", "new_name": "reise"}, cats
    )
    assert cats.names() == ["Internal", "Reise"], cats.names()
    assert "already exists" in out, out
    log("  collision refused, list untouched")


def test_rename_category_reports_a_missing_name():
    cats = FakeCategories(["Internal"])
    out = call_cat_tool("rename_category", {"name": "Nope", "new_name": "X"}, cats)
    assert "no category named" in out, out
    assert cats.names() == ["Internal"]
    log("  missing source name reported plainly")


def test_delete_category_removes_by_name():
    cats = FakeCategories(["Internal", "Reise", "SMA"])
    out = call_cat_tool("delete_category", {"name": "reise"}, cats)
    assert cats.names() == ["Internal", "SMA"], cats.names()
    assert "keep the name" in out, out
    log(f"  deleted -> {cats.names()}, item-side consequence stated")


def test_delete_category_reports_a_missing_name():
    cats = FakeCategories(["Internal"])
    out = call_cat_tool("delete_category", {"name": "Nope"}, cats)
    assert "no category named" in out, out
    assert cats.names() == ["Internal"]
    log("  deleting something absent changes nothing")


def test_color_map_is_contiguous_and_reversible():
    """Guard for the typelib-derived palette: 0..25, no duplicate names."""
    from outlook_desktop_mcp.tools._folder_constants import (
        CATEGORY_COLOR_FROM_NAME,
        CATEGORY_COLOR_NAMES,
    )
    assert sorted(CATEGORY_COLOR_NAMES) == list(range(26)), "gaps in the enum"
    assert len(CATEGORY_COLOR_FROM_NAME) == 26, "duplicate color names"
    assert CATEGORY_COLOR_FROM_NAME["blue"] == 8
    assert CATEGORY_COLOR_NAMES[25] == "dark_maroon"
    log("  26 colors, 0..25 contiguous, name<->index reversible")


def test_set_category_color_changes_it_in_place():
    """Measured: Color is writable, so no delete-and-recreate is needed."""
    cats = FakeCategories(["Internal", "Absence"])
    cats.Item(2).Color = 13                       # gray
    out = call_cat_tool(
        "set_category_color", {"name": "absence", "color": "black"}, cats
    )
    assert cats.Item(2).Color == 15, cats.Item(2).Color
    assert cats.names() == ["Internal", "Absence"], cats.names()
    assert "gray" in out and "black" in out, out
    log(f"  recolored in place, name and position unchanged: {out}")


def test_set_category_color_rejects_an_unknown_color():
    cats = FakeCategories(["Internal"])
    before = cats.Item(1).Color
    out = call_cat_tool(
        "set_category_color", {"name": "Internal", "color": "chartreuse"}, cats
    )
    assert cats.Item(1).Color == before, cats.Item(1).Color
    assert "chartreuse" in out, out
    log("  invalid color leaves the entry untouched")


def test_set_category_color_reports_a_missing_name():
    cats = FakeCategories(["Internal"])
    out = call_cat_tool("set_category_color", {"name": "Nope", "color": "red"}, cats)
    assert "no category named" in out, out
    log("  missing name reported plainly")


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
        ("Create adds with named color", test_create_category_adds_with_named_color),
        ("Create is idempotent", test_create_category_is_idempotent_not_an_error),
        ("Create rejects unknown color", test_create_category_rejects_an_unknown_color),
        ("Rename touches only the list", test_rename_category_renames_only_the_list_entry),
        ("Rename refuses a collision", test_rename_category_refuses_a_collision),
        ("Rename reports missing name", test_rename_category_reports_a_missing_name),
        ("Delete removes by name", test_delete_category_removes_by_name),
        ("Delete reports missing name", test_delete_category_reports_a_missing_name),
        ("Color map contiguous/reversible", test_color_map_is_contiguous_and_reversible),
        ("Recolor in place", test_set_category_color_changes_it_in_place),
        ("Recolor rejects unknown color", test_set_category_color_rejects_an_unknown_color),
        ("Recolor reports missing name", test_set_category_color_reports_a_missing_name),
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
