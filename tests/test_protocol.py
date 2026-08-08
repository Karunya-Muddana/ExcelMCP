"""Regression test: the MCP stdio channel must carry nothing but JSON-RPC.

The server is spawned as a real subprocess and driven through a full
initialize → notifications/initialized → tools/call conversation. Every line
it writes to stdout must parse as JSON. Diagnostics belong on stderr; a single
stray print() on stdout desyncs strict MCP clients, so this test exists to
catch the next debug print before it ships.
"""

import json
import os
import subprocess
import sys


def _frame(obj: dict) -> bytes:
    return (json.dumps(obj) + "\n").encode("utf-8")


def test_stdout_is_pure_jsonrpc(tmp_path):
    env = dict(os.environ)
    # Point ~ at tmp_path so the subprocess can never read or write the real
    # ~/.excelmcp, and so the query call takes the "no index found" diagnostic
    # path — exactly the one that used to print to stdout.
    env["USERPROFILE"] = str(tmp_path)
    env["HOME"] = str(tmp_path)
    env.pop("EXCELMCP_DEFAULT_FOLDER", None)

    requests = [
        {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2025-06-18",
                "capabilities": {},
                "clientInfo": {"name": "protocol-test", "version": "0"},
            },
        },
        {"jsonrpc": "2.0", "method": "notifications/initialized"},
        {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/call",
            "params": {
                "name": "query",
                "arguments": {"question": "total revenue", "folder_path": "/ERP"},
            },
        },
        {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {
                "name": "get_workspace_graph",
                "arguments": {"folder_path": "/ERP"},
            },
        },
    ]

    proc = subprocess.Popen(
        [sys.executable, "-m", "excelmcp.main"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
    )
    stdout, stderr = proc.communicate(
        b"".join(_frame(r) for r in requests), timeout=120
    )

    out_lines = [
        line for line in stdout.decode("utf-8", "replace").splitlines() if line.strip()
    ]
    assert out_lines, "server wrote nothing to stdout — did it start?"

    parsed = []
    bad_lines = []
    for line in out_lines:
        try:
            parsed.append(json.loads(line))
        except json.JSONDecodeError:
            bad_lines.append(line)
    assert not bad_lines, (
        "Non-JSON lines on the MCP stdout channel (move diagnostics to "
        f"storage.log/stderr): {bad_lines!r}"
    )

    # All three requests must have been answered — a desynced or crashed
    # server that emitted only valid JSON would still fail here.
    answered = {m.get("id") for m in parsed if isinstance(m, dict)}
    assert {1, 2, 3} <= answered, (
        f"unanswered requests; got ids {answered}. "
        f"Server stderr:\n{stderr.decode('utf-8', 'replace')}"
    )

    # The diagnostic the query path emits with an empty index must have gone
    # to stderr, proving the path that used to pollute stdout was exercised.
    assert "[Embed]" in stderr.decode("utf-8", "replace")
