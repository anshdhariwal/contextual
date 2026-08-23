#!/usr/bin/env node
const { spawn } = require("child_process");
const path = require("path");

const script = path.join(__dirname, "..", "contextual.py");

const candidates = process.platform === "win32"
  ? [["python"], ["py", "-3"]]
  : [["python3"], ["python"]];

function tryNext() {
  const next = candidates.shift();
  if (!next) {
    console.error("\n  python 3.10+ is required but was not found on PATH.");
    console.error("  grab it from https://www.python.org/downloads/ and run again.\n");
    process.exit(1);
  }
  const child = spawn(next[0], next.slice(1).concat(script), {
    stdio: "inherit",
    cwd: process.cwd(),
  });
  child.on("error", tryNext);
  child.on("exit", (code) => { process.exitCode = code === null ? 1 : code; });
}

tryNext();
