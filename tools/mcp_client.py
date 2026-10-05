"""Drive the ferroxl-mcp server over stdio JSON-RPC.

A client that keeps one subprocess alive, so a test can make a sequence of calls against
one workbook the way an agent would. Everything here is protocol-level: no Rust is
imported and nothing is stubbed, so what passes here is what an agent would actually see.
"""

import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def find_binary() -> Path:
    """Locate the built server, honouring CARGO_TARGET_DIR the way cargo does."""
    target = Path(os.environ.get("CARGO_TARGET_DIR") or (ROOT.parent / "target"))
    name = "ferroxl-mcp.exe" if os.name == "nt" else "ferroxl-mcp"
    for profile in ("debug", "release"):
        candidate = target / profile / name
        if candidate.is_file():
            return candidate
    raise SystemExit(
        f"ferroxl-mcp not built under {target}; run `cargo build -p ferroxl-mcp` first"
    )


class Server:
    """One server process, spoken to over line-delimited JSON-RPC."""

    def __init__(self, root: Path):
        self.root = root
        self.next_id = 0
        self.proc = subprocess.Popen(
            [str(find_binary()), "--root", str(root)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            bufsize=1,
        )

    def send(self, method: str, params: dict | None = None) -> dict:
        self.next_id += 1
        message = {"jsonrpc": "2.0", "id": self.next_id, "method": method}
        if params is not None:
            message["params"] = params
        self.proc.stdin.write(json.dumps(message) + "\n")
        self.proc.stdin.flush()
        line = self.proc.stdout.readline()
        if not line:
            stderr = self.proc.stderr.read()
            raise RuntimeError(f"server closed the stream\n{stderr}")
        return json.loads(line)

    def initialize(self) -> dict:
        return self.send(
            "initialize",
            {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "parity", "version": "0"},
            },
        )

    def list_tools(self) -> list[dict]:
        return self.send("tools/list")["result"]["tools"]

    def call(self, tool: str, **arguments) -> dict:
        """Return the whole result envelope, because `isError` is part of the contract.

        The first parameter is `tool` rather than `name` because several tools take an
        argument called `name`, and a positional `name` would collide with it.
        """
        response = self.send("tools/call", {"name": tool, "arguments": arguments})
        if "error" in response:
            return {"rpcError": response["error"]}
        return response["result"]

    def ok(self, tool: str, **arguments) -> dict:
        """Call a tool and fail loudly if it reported an error."""
        result = self.call(tool, **arguments)
        if result.get("isError"):
            text = " ".join(part.get("text", "") for part in result.get("content", []))
            raise AssertionError(f"{tool} returned isError: {text}")
        if "rpcError" in result:
            raise AssertionError(f"{tool} returned a JSON-RPC error: {result['rpcError']}")
        return result

    def data(self, tool: str, **arguments):
        """The `structuredContent` of a successful call."""
        return self.ok(tool, **arguments).get("structuredContent")

    def summary(self, tool: str, **arguments) -> str:
        """The one-line human-readable summary, which is what a model reads first."""
        parts = self.ok(tool, **arguments).get("content", [])
        return " ".join(part.get("text", "") for part in parts)

    def close(self) -> None:
        try:
            self.proc.stdin.close()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=10)
        except Exception:
            self.proc.kill()

    def __enter__(self):
        return self

    def __exit__(self, *exception):
        self.close()


if __name__ == "__main__":
    # A smoke test: the server answers `initialize` and advertises its catalogue.
    with Server(Path(sys.argv[1] if len(sys.argv) > 1 else ROOT / "target" / "probe")) as s:
        s.initialize()
        tools = s.list_tools()
        print(f"{len(tools)} tools")
        for tool in tools:
            print(f"  {tool['name']}")