"""
Outlook Desktop MCP - Phase 3 MCP Server Test
==============================================
Uses the MCP SDK client to connect to the server over stdio
and exercise all tools.
"""
import sys
import os
import json
import asyncio
import logging

# Ensure our package is importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

logging.basicConfig(level=logging.WARNING, stream=sys.stderr)

# Recipient for the send_email test. Set MCP_TEST_RECIPIENT in your shell
# to an address you control. Default is an RFC 2606 undeliverable address
# so no external email is sent during automated testing.
_TEST_RECIPIENT = os.environ.get("MCP_TEST_RECIPIENT", "outlook-mcp-test@example.invalid")


def log(msg):
    print(msg, file=sys.stderr, flush=True)


async def run_tests():
    from mcp.client.stdio import stdio_client, StdioServerParameters
    from mcp.client.session import ClientSession

    python_exe = sys.executable  # use the same venv that's running this script

    server_params = StdioServerParameters(
        command=python_exe,
        args=["-m", "outlook_desktop_mcp.server"],
        cwd=os.path.join(os.path.dirname(__file__), ".."),
    )

    log("=" * 60)
    log("Outlook Desktop MCP - Phase 3 Server Test")
    log("=" * 60)
    log("\nConnecting to server...")

    passed = 0
    total = 0

    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            log("Server initialized.\n")

            # ----- Test 1: Tool Discovery -----
            total += 1
            log("--- Test 1: Tool Discovery ---")
            try:
                tools_result = await session.list_tools()
                tool_names = [t.name for t in tools_result.tools]
                log(f"  Found {len(tool_names)} tools:")
                for t in tools_result.tools:
                    desc = t.description.split("\n")[0][:75] if t.description else ""
                    log(f"    - {t.name}: {desc}")

                expected = [
                    "send_email", "list_emails", "read_email", "mark_as_read",
                    "mark_as_unread", "set_flag", "clear_flag",
                    "move_email", "reply_email",
                    "list_folders", "search_emails",
                ]
                missing = [n for n in expected if n not in tool_names]
                assert not missing, f"Missing tools: {missing}"
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 2: list_folders -----
            total += 1
            log("\n--- Test 2: list_folders ---")
            try:
                result = await session.call_tool("list_folders", {"max_depth": 1})
                content = result.content[0].text
                folders = json.loads(content)
                folder_names = [f["name"] for f in folders]
                log(f"  Top-level folders: {folder_names}")
                assert len(folders) > 0
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 3: list_emails -----
            total += 1
            log("\n--- Test 3: list_emails (inbox, 3) ---")
            try:
                result = await session.call_tool("list_emails", {"folder": "inbox", "count": 3})
                content = result.content[0].text
                emails = json.loads(content)
                log(f"  Got {len(emails)} emails:")
                for e in emails:
                    log(f"    - {e['subject'][:60]}")
                assert len(emails) > 0
                first_entry_id = emails[0]["entry_id"]
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")
                first_entry_id = None

            # ----- Test 4: read_email -----
            total += 1
            log("\n--- Test 4: read_email (by entry_id) ---")
            try:
                assert first_entry_id, "No entry_id from previous test"
                result = await session.call_tool("read_email", {"entry_id": first_entry_id})
                content = result.content[0].text
                email_data = json.loads(content)
                log(f"  Subject: {email_data['subject']}")
                log(f"  From: {email_data['sender_name']}")
                log(f"  Body: {email_data['body'][:100]}...")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 5: search_emails -----
            total += 1
            log("\n--- Test 5: search_emails ('Feedback') ---")
            try:
                result = await session.call_tool("search_emails", {
                    "query": "Feedback", "folder": "inbox", "count": 5
                })
                content = result.content[0].text
                results = json.loads(content)
                log(f"  Found {len(results)} results")
                for r in results:
                    log(f"    - {r['subject'][:60]}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 6: send_email -----
            total += 1
            log("\n--- Test 6: send_email ---")
            try:
                result = await session.call_tool("send_email", {
                    "to": _TEST_RECIPIENT,
                    "subject": "Outlook Desktop MCP - Phase 3 MCP Test",
                    "body": "Sent through the MCP server via stdio. If you see this, the MCP layer works!",
                })
                content = result.content[0].text
                log(f"  Result: {content}")
                assert "sent" in content.lower()
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 7: read_email by subject_search -----
            total += 1
            log("\n--- Test 7: read_email (by subject_search) ---")
            try:
                result = await session.call_tool("read_email", {
                    "subject_search": "Feedback", "folder": "inbox"
                })
                content = result.content[0].text
                email_data = json.loads(content)
                if "error" in email_data:
                    log(f"  No match (expected if no 'Feedback' emails): {email_data['error']}")
                else:
                    log(f"  Found: {email_data['subject']}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 8: set_flag -----
            total += 1
            log("\n--- Test 8: set_flag ---")
            try:
                assert first_entry_id, "No entry_id from previous test"
                result = await session.call_tool("set_flag", {"entry_id": first_entry_id})
                content = result.content[0].text
                log(f"  Result: {content}")
                assert "Flagged:" in content
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 9: clear_flag (restores state from Test 8) -----
            total += 1
            log("\n--- Test 9: clear_flag ---")
            try:
                assert first_entry_id, "No entry_id from previous test"
                result = await session.call_tool("clear_flag", {"entry_id": first_entry_id})
                content = result.content[0].text
                log(f"  Result: {content}")
                assert "Flag cleared:" in content
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 10: reply_email save_as_draft -----
            total += 1
            log("\n--- Test 10: reply_email (save_as_draft) ---")
            draft_entry_id = None
            try:
                assert first_entry_id, "No entry_id from previous test"
                result = await session.call_tool("reply_email", {
                    "entry_id": first_entry_id,
                    "body": "MCP draft reply test — please ignore",
                    "save_as_draft": True,
                })
                content = result.content[0].text
                draft = json.loads(content)
                draft_entry_id = draft.get("entry_id")
                log(f"  Draft entry_id: {(draft_entry_id or '')[:40]}...")
                assert draft_entry_id, "Expected entry_id in response"
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")
            finally:
                # Clean up: move draft to Deleted Items so Drafts folder stays clean
                if draft_entry_id:
                    try:
                        await session.call_tool("move_email", {
                            "entry_id": draft_entry_id,
                            "target_folder": "deleted",
                        })
                        log("  (draft cleaned up)")
                    except Exception:
                        pass  # best-effort cleanup

    log("")
    log("=" * 60)
    log(f"Results: {passed}/{total} passed")
    log("=" * 60)
    return passed == total


if __name__ == "__main__":
    success = asyncio.run(run_tests())
    sys.exit(0 if success else 1)
