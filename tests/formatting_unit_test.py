"""
Outlook Desktop MCP - Formatting Unit Tests
=============================================
Pure-logic tests for field selection and the summary envelope. Needs NO COM
and no running Outlook, so it runs in CI and on any machine.

Run: .venv\\Scripts\\python tests\\formatting_unit_test.py
"""
import sys

from outlook_desktop_mcp.utils.formatting import (
    EMAIL_SUMMARY_FIELDS,
    format_email_summary,
    parse_summary_fields,
)


def log(msg):
    print(msg, file=sys.stderr, flush=True)


class FakeAttachments:
    def __init__(self, count, counter):
        self._count = count
        self._counter = counter

    @property
    def Count(self):
        self._counter["attachments"] += 1
        return self._count


class FakeMail:
    """Records which properties were touched, so we can assert that an
    unrequested field costs no COM access."""

    def __init__(self, attachment_count=2):
        self.touched = []
        self.counter = {"attachments": 0}
        self._attachments = FakeAttachments(attachment_count, self.counter)

    def _mark(self, name, value):
        self.touched.append(name)
        return value

    @property
    def EntryID(self):
        return self._mark("EntryID", "0000ABCD")

    @property
    def Subject(self):
        return self._mark("Subject", "Quarterly report")

    @property
    def SenderEmailAddress(self):
        return self._mark("SenderEmailAddress", "someone@example.com")

    @property
    def SenderName(self):
        return self._mark("SenderName", "Someone")

    @property
    def ReceivedTime(self):
        return self._mark("ReceivedTime", "2026-08-19 21:39:46")

    @property
    def UnRead(self):
        return self._mark("UnRead", True)

    @property
    def FlagStatus(self):
        return self._mark("FlagStatus", 0)

    @property
    def Attachments(self):
        return self._attachments


def test_default_returns_all_fields():
    """No field spec keeps the historical shape, in the declared order."""
    out = format_email_summary(FakeMail())
    assert tuple(out.keys()) == EMAIL_SUMMARY_FIELDS, out.keys()
    log(f"  {len(out)} fields, order preserved")


def test_subset_selects_and_orders():
    """A subset returns only those keys, still in declared order."""
    out = format_email_summary(FakeMail(), ("received_time", "subject"))
    assert list(out.keys()) == ["subject", "received_time"], out.keys()
    assert out["subject"] == "Quarterly report"
    log(f"  {out}")


def test_unrequested_fields_are_not_read():
    """The point of the feature: skipped fields cost no COM round trip."""
    m = FakeMail()
    format_email_summary(m, ("subject",))
    assert m.touched == ["Subject"], m.touched
    assert m.counter["attachments"] == 0, m.counter
    log(f"  touched only {m.touched}, Attachments.Count calls: {m.counter['attachments']}")


def test_attachment_fields_share_one_com_call():
    """has_attachments + attachment_count must not fetch Count twice."""
    m = FakeMail(attachment_count=3)
    out = format_email_summary(m, ("has_attachments", "attachment_count"))
    assert out == {"has_attachments": True, "attachment_count": 3}, out
    assert m.counter["attachments"] == 1, m.counter
    log(f"  Attachments.Count calls: {m.counter['attachments']} (was 2 before)")


def test_parse_empty_means_all():
    for spec in ("", "   ", ","):
        fields, err = parse_summary_fields(spec)
        assert fields is None and err is None, (spec, fields, err)
    log("  empty / blank / bare-comma all mean 'all fields'")


def test_parse_valid_spec():
    fields, err = parse_summary_fields(" subject , received_time ")
    assert err is None, err
    assert fields == ("subject", "received_time"), fields
    log(f"  parsed {fields}")


def test_parse_deduplicates():
    fields, err = parse_summary_fields("subject,subject,unread")
    assert err is None, err
    assert fields == ("subject", "unread"), fields
    log(f"  deduplicated to {fields}")


def test_parse_rejects_unknown():
    """A typo must fail loudly, not silently drop the field."""
    fields, err = parse_summary_fields("subject,sendername")
    assert fields is None, fields
    assert err and "sendername" in err, err
    assert "sender_name" in err, "error should list the valid names"
    log(f"  rejected: {err[:60]}...")


def test_envelope_flags_truncation():
    from outlook_desktop_mcp.server import _summary_envelope
    rows = [{"received_time": "2026-08-20 10:00"}, {"received_time": "2026-08-19 21:39"}]
    env = _summary_envelope(rows, total=250)
    assert env["returned"] == 2
    assert env["total_matching"] == 250
    assert env["truncated"] is True
    assert env["oldest_returned"] == "2026-08-19 21:39"
    log(f"  truncated={env['truncated']} cursor={env['oldest_returned']}")


def test_envelope_complete_and_empty():
    from outlook_desktop_mcp.server import _summary_envelope
    env = _summary_envelope([{"received_time": "x"}], total=1)
    assert env["truncated"] is False, env
    empty = _summary_envelope([], total=0)
    assert empty["truncated"] is False and empty["oldest_returned"] == "", empty
    log("  complete list and empty list both report truncated=False")


def test_envelope_cursor_empty_without_received_time():
    """Documented edge: excluding received_time removes the pagination cursor."""
    from outlook_desktop_mcp.server import _summary_envelope
    env = _summary_envelope([{"subject": "no timestamp here"}], total=99)
    assert env["truncated"] is True
    assert env["oldest_returned"] == "", env
    log("  cursor is empty when received_time was not selected")


def test_no_hardcoded_us_date_format_in_filters():
    """Regression guard for the de-DE month/day swap in the MAIL date filters.

    Outlook's Restrict() and its DASL variant parse date strings with the
    Windows system locale. list_emails / search_emails formatted theirs as
    '%m/%d/%Y', so on a de-DE box '2026-09-11' went over the wire as
    '09/11/2026' and came back read as 9 November. Measured 2026-09-11:
    start_date='2026-09-11' returned 0 mails while start_date='2026-11-09'
    returned that day's 10 -- same tool, same day, inverted spelling.

    The bug only shows when the day is <= 12; above that there is no such
    month and Outlook falls back to the right reading, which is why earlier
    checks on e.g. 2026-08-17 kept declaring it fixed. A value-level test
    would have to pin a locale to mean anything, so this guards the source
    instead: every date handed to a filter must go through
    _date_to_restrict_str, which formats in the locale's own order.
    """
    from pathlib import Path
    server_py = (
        Path(__file__).resolve().parent.parent
        / "src" / "outlook_desktop_mcp" / "server.py"
    )
    source = server_py.read_text(encoding="utf-8")
    assert "%m/%d/%Y" not in source, (
        "server.py hardcodes the US date order somewhere. Outlook parses "
        "Restrict/DASL dates by system locale -- use _date_to_restrict_str()."
    )
    log("  no hardcoded US date order in server.py")


def main():
    tests = [
        ("No hardcoded US date format", test_no_hardcoded_us_date_format_in_filters),
        ("Default returns all fields", test_default_returns_all_fields),
        ("Subset selects and orders", test_subset_selects_and_orders),
        ("Unrequested fields are not read", test_unrequested_fields_are_not_read),
        ("Attachment fields share one COM call", test_attachment_fields_share_one_com_call),
        ("Empty spec means all", test_parse_empty_means_all),
        ("Valid spec parses", test_parse_valid_spec),
        ("Spec deduplicates", test_parse_deduplicates),
        ("Unknown field rejected", test_parse_rejects_unknown),
        ("Envelope flags truncation", test_envelope_flags_truncation),
        ("Envelope complete/empty", test_envelope_complete_and_empty),
        ("Cursor empty without received_time", test_envelope_cursor_empty_without_received_time),
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
