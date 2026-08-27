#!/usr/bin/env python3
"""Static checks for exported VBA source (.bas/.cls/.frm). Pure stdlib.

Usage:
  python3 tools/vba_lint.py [paths...] [--json] [--max-severity warn]

Defaults to linting src/ and tests/. Exit code 1 if any finding is at or above
the failing severity (default: error), so it works as a CI/agent gate.

Rules are deliberately mechanical - they catch the VBA mistakes that are cheap
to spot and expensive to debug. Judgement calls stay with the review agents.
"""
import argparse
import json
import re
import sys
from pathlib import Path

SEVERITY_ORDER = {"info": 0, "warn": 1, "error": 2}

PROC_START = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)",
    re.I,
)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
DIM_NO_TYPE = re.compile(r"^\s*(?:Dim|Static)\s+([A-Za-z_]\w*)\s*(?:,|$)", re.I)
UNQUALIFIED = re.compile(r"(?<![.\w])(Range|Cells|Rows|Columns)\s*\(", re.I)
MAGIC_SHEET = re.compile(r"(?:Worksheets|Sheets)\s*\(\s*\d+\s*\)", re.I)
HARDCODED_PATH = re.compile(r'"(?:[A-Za-z]:\\[^"]+|\\\\[A-Za-z0-9_.$-]+\\[^"]*)"')
RISKY = {
    "SendKeys": ("warn", "SendKeys is timing-dependent and silently wrong when focus moves"),
    "Shell": ("warn", "Shell runs external commands - validate inputs and document why"),
    "Kill": ("warn", "Kill deletes files with no undo - guard with an existence check and a confirmation"),
    "WScript.Shell": ("warn", "CreateObject(\"WScript.Shell\") is a common macro-malware pattern; reviewers will flag it"),
    "Auto_Open": ("info", "Auto_Open runs on open - keep it thin and guarded"),
}
STATE_TOGGLES = {
    "screenupdating": "Application.ScreenUpdating",
    "enableevents": "Application.EnableEvents",
    "displayalerts": "Application.DisplayAlerts",
    "calculation": "Application.Calculation",
}
LONG_PROC_LINES = 80


def strip_noise(line):
    """Blank out string literals and drop trailing comments so regexes see code only."""
    out, in_str = [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
            out.append('"')
        elif in_str:
            out.append(" ")
        elif ch == "'":
            break
        else:
            out.append(ch)
    text = "".join(out)
    return "" if re.match(r"^\s*Rem\b", text, re.I) else text


class Finding(dict):
    def __init__(self, file, line, rule, severity, message, hint=""):
        super().__init__(
            file=str(file), line=line, rule=rule, severity=severity, message=message, hint=hint
        )


def lint_file(path):
    raw = path.read_text(encoding="utf-8", errors="replace").splitlines()
    code = [strip_noise(l) for l in raw]
    findings = []
    add = findings.append

    head = "\n".join(code[:20]).lower()
    if "option explicit" not in head and path.suffix.lower() in (".bas", ".cls"):
        add(Finding(path, 1, "option-explicit", "error",
                    "Module has no Option Explicit",
                    "Add 'Option Explicit' as the first line; typos become compile errors instead of empty Variants."))

    proc = None  # name, start_line, has_error_handler, toggles{}
    for idx, line in enumerate(code, start=1):
        low = line.lower()
        m = PROC_START.match(line)
        if m and proc is None:
            proc = {"name": m.group(2), "start": idx, "err": False, "toggles": {}, "kind": m.group(1)}
        if PROC_END.match(line) and proc:
            length = idx - proc["start"]
            if not proc["err"] and length > 12:
                add(Finding(path, proc["start"], "no-error-handler", "warn",
                            f"{proc['kind']} {proc['name']} has no error handling",
                            "Add 'On Error GoTo Fail' plus a Fail: block, or state in a comment why failure is impossible."))
            if length > LONG_PROC_LINES:
                add(Finding(path, proc["start"], "long-procedure", "info",
                            f"{proc['kind']} {proc['name']} is {length} lines",
                            "Split into named steps so each piece can be tested on its own."))
            for key, restored in proc["toggles"].items():
                if not restored:
                    add(Finding(path, proc["start"], "unrestored-app-state", "error",
                                f"{proc['name']} sets {STATE_TOGGLES[key]} but never restores it",
                                "Restore it in the error handler too, not only on the happy path - otherwise Excel is left in that state."))
            proc = None

        if re.search(r"On\s+Error\s+GoTo\s+(?!0\b)", line, re.I):
            if proc:
                proc["err"] = True
        if re.search(r"On\s+Error\s+Resume\s+Next", line, re.I):
            window = " ".join(code[idx: idx + 15]).lower()
            if "on error goto 0" not in window:
                add(Finding(path, idx, "resume-next-unscoped", "error",
                            "On Error Resume Next is not closed with On Error GoTo 0 nearby",
                            "Scope it to the single statement that can fail, then restore error handling immediately."))
            if proc:
                proc["err"] = True

        for key, label in STATE_TOGGLES.items():
            if re.search(rf"Application\s*\.\s*{key}\s*=", low, re.I) and proc:
                value = low.split("=", 1)[1].strip()
                turning_off = value in ("false", "xlcalculationmanual")
                if turning_off:
                    proc["toggles"].setdefault(key, False)
                else:
                    proc["toggles"][key] = True

        if re.search(r"(?<![.\w])(Selection|ActiveCell|ActiveSheet|ActiveWorkbook)\b", line, re.I) or \
           re.search(r"\.\s*(Select|Activate)\s*$", line, re.I):
            add(Finding(path, idx, "select-activate", "warn",
                        "Depends on the active selection",
                        "Work against an explicit Worksheet/Range object - selection-based code breaks when the user clicks elsewhere."))
        if UNQUALIFIED.search(line) and not re.match(r"^\s*(?:'|Attribute)", line):
            add(Finding(path, idx, "unqualified-range", "warn",
                        "Range/Cells call is not qualified with a worksheet",
                        "Qualify it (ws.Range(...)) - unqualified calls silently target whichever sheet is active."))
        if MAGIC_SHEET.search(line):
            add(Finding(path, idx, "sheet-by-index", "warn",
                        "Sheet referenced by index",
                        "Use the sheet's CodeName or a defined name; indexes change when users reorder tabs."))
        dm = DIM_NO_TYPE.match(line)
        if dm and " as " not in low:
            add(Finding(path, idx, "untyped-variable", "warn",
                        f"'{dm.group(1)}' declared without a type",
                        "Give it an explicit type - untyped means Variant, which hides type errors until runtime."))
        if re.search(r"\bAs\s+Integer\b", line, re.I):
            add(Finding(path, idx, "integer-type", "info",
                        "As Integer overflows above 32767",
                        "Use Long for anything counting rows or loop iterations."))
        if HARDCODED_PATH.search(raw[idx - 1]):
            add(Finding(path, idx, "hardcoded-path", "warn",
                        "Hardcoded file path",
                        "Take the path from a parameter, a config sheet, or ThisWorkbook.Path."))
        if re.match(r"^\s*(Public|Global)\s+(?!(?:Sub|Function|Property|Const|Enum|Type|Declare))", line, re.I):
            add(Finding(path, idx, "public-module-state", "info",
                        "Module-level Public variable",
                        "Prefer passing state as arguments; globals make behaviour depend on execution order."))
        for token, (sev, msg) in RISKY.items():
            if token.lower() in low:
                add(Finding(path, idx, "risky-call", sev, msg,
                            "Keep it, but document the intent and validate any user-supplied input."))
        if re.search(r"\.\s*Value\b(?!2)", line) and ".value2" not in low:
            add(Finding(path, idx, "value-vs-value2", "info",
                        "Uses .Value rather than .Value2",
                        ".Value2 is faster and avoids Currency/Date coercion surprises - unless you want dates as dates."))
    return findings


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("paths", nargs="*", default=None)
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    ap.add_argument("--fail-on", default="error", choices=list(SEVERITY_ORDER),
                    help="lowest severity that makes this exit non-zero")
    args = ap.parse_args(argv)

    roots = [Path(p) for p in (args.paths or ["src", "tests"])]
    files = []
    for root in roots:
        if root.is_dir():
            files += sorted(
                f for ext in ("*.bas", "*.cls", "*.frm") for f in root.rglob(ext)
            )
        elif root.is_file():
            files.append(root)
    findings = [f for path in files for f in sorted(lint_file(path), key=lambda f: f["line"])]
    threshold = SEVERITY_ORDER[args.fail_on]
    failing = [f for f in findings if SEVERITY_ORDER[f["severity"]] >= threshold]

    if args.json:
        print(json.dumps({"files_linted": [str(f) for f in files], "findings": findings,
                          "failing": len(failing)}, indent=2))
    else:
        for f in findings:
            print(f"{f['file']}:{f['line']}: {f['severity']}: [{f['rule']}] {f['message']}")
            if f["hint"]:
                print(f"    -> {f['hint']}")
        print(f"\n{len(files)} file(s), {len(findings)} finding(s), {len(failing)} at or above '{args.fail_on}'")
    return 1 if failing else 0


if __name__ == "__main__":
    sys.exit(main())
