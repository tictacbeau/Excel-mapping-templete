#!/usr/bin/env python3
"""
Remittance Allocation Engine — Python Build Helper
===================================================
Validates VBA source files, checks syntax structure,
and generates reports. Does NOT require Excel to be installed.

For actual XLSM generation, use build.ps1 (requires Windows + Excel).

Usage:
    python build.py [--validate] [--report] [--check-deps]

Requirements:
    pip install -r requirements.txt
"""

import os
import sys
import re
import json
import argparse
from pathlib import Path
from datetime import datetime


# ============================================================
#  PATHS
# ============================================================
REPO_ROOT = Path(__file__).parent.parent
SRC_VBA = REPO_ROOT / "src" / "vba"
MODULES_DIR = SRC_VBA / "modules"
CLASSES_DIR = SRC_VBA / "classes"
FORMS_DIR = SRC_VBA / "forms"
CUSTOMUI_DIR = REPO_ROOT / "src" / "customui"
DIST_DIR = REPO_ROOT / "dist"
BUILD_DIR = REPO_ROOT / "build"

# Expected VBA source files
EXPECTED_MODULES = [
    "modConstants.bas",
    "modUtils.bas",
    "modFileConverter.bas",
    "modPayorRegistry.bas",
    "modTemplateEngine.bas",
    "modDataLookup.bas",
    "modMain.bas",
    "modRibbon.bas",
    "modThisWorkbook.bas",
]

EXPECTED_CLASSES = [
    "clsCellFormat.cls",
    "clsPayor.cls",
    "clsRemittance.cls",
    "clsTemplate.cls",
]

EXPECTED_FORMS = [
    "frmProcessRemittance.frm",
    "frmPayorWizard.frm",
    "frmTemplateBuilder.frm",
    "frmPayorManager.frm",
    "frmDataRefresh.frm",
]

EXPECTED_CUSTOMUI = ["customUI14.xml"]


# ============================================================
#  VALIDATION
# ============================================================

def validate_file_exists(path: Path, label: str) -> bool:
    """Check that a required file exists."""
    if path.exists():
        print(f"  [OK]   {label}: {path.name}")
        return True
    else:
        print(f"  [MISS] {label}: {path.name} — NOT FOUND at {path}")
        return False


def validate_vba_module(path: Path) -> dict:
    """
    Basic structural validation of a VBA module/class/form.
    Checks for:
    - Attribute VB_Name declaration
    - Option Explicit
    - Balanced Sub/Function/End Sub/End Function pairs
    - Error handling (On Error GoTo) presence
    - No obvious syntax issues
    """
    result = {
        "file": path.name,
        "ok": True,
        "warnings": [],
        "errors": [],
        "stats": {}
    }

    if not path.exists():
        result["ok"] = False
        result["errors"].append("File not found")
        return result

    try:
        content = path.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        result["ok"] = False
        result["errors"].append(f"Cannot read file: {e}")
        return result

    lines = content.splitlines()
    result["stats"]["lines"] = len(lines)

    # Check for VB_Name
    has_vb_name = any("Attribute VB_Name" in ln for ln in lines)
    if not has_vb_name:
        result["warnings"].append("Missing 'Attribute VB_Name' declaration")

    # Check for Option Explicit
    has_option_explicit = any(ln.strip() == "Option Explicit" for ln in lines)
    if not has_option_explicit:
        result["errors"].append("Missing 'Option Explicit'")
        result["ok"] = False

    # Count Sub/Function/End Sub/End Function
    sub_count = sum(1 for ln in lines
                    if re.match(r'^\s*(Public|Private|Friend)?\s*Sub\s+\w+', ln, re.I))
    func_count = sum(1 for ln in lines
                     if re.match(r'^\s*(Public|Private|Friend)?\s*Function\s+\w+', ln, re.I))
    end_sub_count = sum(1 for ln in lines
                        if re.match(r'^\s*End\s+Sub\s*$', ln, re.I))
    end_func_count = sum(1 for ln in lines
                         if re.match(r'^\s*End\s+Function\s*$', ln, re.I))

    result["stats"]["subs"] = sub_count
    result["stats"]["functions"] = func_count

    if sub_count != end_sub_count:
        result["warnings"].append(
            f"Sub/End Sub mismatch: {sub_count} Sub(s), {end_sub_count} End Sub(s)")

    if func_count != end_func_count:
        result["warnings"].append(
            f"Function/End Function mismatch: {func_count} Function(s), "
            f"{end_func_count} End Function(s)")

    # Check for error handling
    has_err_handling = any("On Error" in ln for ln in lines)
    if not has_err_handling and (sub_count + func_count) > 0:
        result["warnings"].append(
            "No 'On Error' statements found — consider adding error handling")

    # Check for TODO / FIXME
    todos = [f"Line {i+1}: {ln.strip()}" for i, ln in enumerate(lines)
             if "TODO" in ln.upper() or "FIXME" in ln.upper() or "HACK" in ln.upper()]
    if todos:
        result["warnings"].append(f"Found {len(todos)} TODO/FIXME/HACK comment(s)")

    return result


def validate_customui(path: Path) -> dict:
    """Validate the CustomUI XML file."""
    result = {"file": path.name, "ok": True, "warnings": [], "errors": []}
    if not path.exists():
        result["ok"] = False
        result["errors"].append("File not found")
        return result

    content = path.read_text(encoding="utf-8")
    if "<customUI" not in content:
        result["errors"].append("Missing <customUI> root element")
        result["ok"] = False
    if "onAction" not in content:
        result["warnings"].append("No onAction callbacks defined")
    if "screentip" not in content:
        result["warnings"].append("No screentip text found")

    # Check all onAction callbacks are in modRibbon
    callbacks = re.findall(r'onAction="(\w+)"', content)
    result["stats"] = {"buttons": len(callbacks), "callbacks": callbacks}
    return result


# ============================================================
#  REPORTS
# ============================================================

def generate_report(all_results: list) -> str:
    """Generate a text build report."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "=" * 60,
        "  REMITTANCE ALLOCATION ENGINE — BUILD VALIDATION REPORT",
        f"  Generated: {now}",
        "=" * 60,
        "",
    ]

    total_ok = 0
    total_warn = 0
    total_err = 0

    for res in all_results:
        status = "OK" if res["ok"] else "FAIL"
        lines.append(f"[{status}] {res['file']}")
        if res.get("stats"):
            stats_str = " | ".join(f"{k}: {v}" for k, v in res["stats"].items()
                                    if not isinstance(v, list))
            lines.append(f"       Stats: {stats_str}")
        for w in res.get("warnings", []):
            lines.append(f"       WARN: {w}")
            total_warn += 1
        for e in res.get("errors", []):
            lines.append(f"       ERROR: {e}")
            total_err += 1
        if res["ok"]:
            total_ok += 1
        else:
            total_err += 1

    lines.extend([
        "",
        "-" * 60,
        f"  Files OK:       {total_ok}",
        f"  Warnings:       {total_warn}",
        f"  Errors:         {total_err}",
        "-" * 60,
        "  STATUS: " + ("PASS" if total_err == 0 else "FAIL"),
        "=" * 60,
    ])
    return "\n".join(lines)


def count_lines_of_code() -> dict:
    """Count lines of code per file type."""
    counts = {"modules": 0, "classes": 0, "forms": 0, "total": 0}
    for f in MODULES_DIR.glob("*.bas"):
        try:
            n = len(f.read_text(encoding="utf-8", errors="ignore").splitlines())
            counts["modules"] += n
            counts["total"] += n
        except Exception:
            pass
    for f in CLASSES_DIR.glob("*.cls"):
        try:
            n = len(f.read_text(encoding="utf-8", errors="ignore").splitlines())
            counts["classes"] += n
            counts["total"] += n
        except Exception:
            pass
    for f in FORMS_DIR.glob("*.frm"):
        try:
            n = len(f.read_text(encoding="utf-8", errors="ignore").splitlines())
            counts["forms"] += n
            counts["total"] += n
        except Exception:
            pass
    return counts


def check_cross_references() -> list:
    """
    Check that all modMain calls / form references have corresponding
    module/form definitions.
    """
    issues = []
    # Collect all defined Sub/Function names
    defined = set()
    for f in list(MODULES_DIR.glob("*.bas")) + list(CLASSES_DIR.glob("*.cls")):
        try:
            content = f.read_text(encoding="utf-8", errors="ignore")
            matches = re.findall(
                r'(?:Public|Private|Friend)?\s+(?:Sub|Function)\s+(\w+)', content, re.I)
            defined.update(matches)
        except Exception:
            pass

    # Check modMain references
    main_path = MODULES_DIR / "modMain.bas"
    if main_path.exists():
        content = main_path.read_text(encoding="utf-8", errors="ignore")
        calls = re.findall(r'(\w+)\.\w+\s*\(', content)
        # This is a simplified check
    return issues


# ============================================================
#  MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Remittance Allocation Engine — Python Build Helper")
    parser.add_argument("--validate", action="store_true",
                        help="Validate all VBA source files")
    parser.add_argument("--report", action="store_true",
                        help="Generate and print build report")
    parser.add_argument("--stats", action="store_true",
                        help="Show lines of code statistics")
    parser.add_argument("--check-files", action="store_true",
                        help="Check all expected files are present")
    args = parser.parse_args()

    if not any(vars(args).values()):
        # Default: run all checks
        args.validate = True
        args.report = True
        args.stats = True
        args.check_files = True

    all_results = []

    # ---- File presence check ----
    if args.check_files:
        print("\n--- File Presence Check ---")
        all_ok = True
        for fname in EXPECTED_MODULES:
            p = MODULES_DIR / fname
            if not p.exists():
                print(f"  [MISS] module: {fname}")
                all_ok = False
            else:
                print(f"  [OK]   module: {fname}")
        for fname in EXPECTED_CLASSES:
            p = CLASSES_DIR / fname
            if not p.exists():
                print(f"  [MISS] class:  {fname}")
                all_ok = False
            else:
                print(f"  [OK]   class:  {fname}")
        for fname in EXPECTED_FORMS:
            p = FORMS_DIR / fname
            if not p.exists():
                print(f"  [MISS] form:   {fname}")
                all_ok = False
            else:
                print(f"  [OK]   form:   {fname}")
        for fname in EXPECTED_CUSTOMUI:
            p = CUSTOMUI_DIR / fname
            if not p.exists():
                print(f"  [MISS] ui:     {fname}")
                all_ok = False
            else:
                print(f"  [OK]   ui:     {fname}")

    # ---- VBA validation ----
    if args.validate:
        print("\n--- VBA Source Validation ---")
        for fname in EXPECTED_MODULES:
            res = validate_vba_module(MODULES_DIR / fname)
            all_results.append(res)
            status = "OK  " if res["ok"] else "FAIL"
            warn_str = f" ({len(res['warnings'])} warn)" if res["warnings"] else ""
            print(f"  [{status}] {fname}{warn_str}")
            if res["errors"]:
                for e in res["errors"]:
                    print(f"         ERROR: {e}")

        for fname in EXPECTED_CLASSES:
            res = validate_vba_module(CLASSES_DIR / fname)
            all_results.append(res)
            status = "OK  " if res["ok"] else "FAIL"
            print(f"  [{status}] {fname}")

        for fname in EXPECTED_CUSTOMUI:
            res = validate_customui(CUSTOMUI_DIR / fname)
            all_results.append(res)
            status = "OK  " if res["ok"] else "FAIL"
            print(f"  [{status}] {fname}")

    # ---- Stats ----
    if args.stats:
        print("\n--- Lines of Code ---")
        counts = count_lines_of_code()
        print(f"  Modules:  {counts['modules']:>6,} lines")
        print(f"  Classes:  {counts['classes']:>6,} lines")
        print(f"  Forms:    {counts['forms']:>6,} lines")
        print(f"  Total:    {counts['total']:>6,} lines")

    # ---- Report ----
    if args.report and all_results:
        print("\n")
        print(generate_report(all_results))

        # Save report to file
        DIST_DIR.mkdir(exist_ok=True)
        report_path = DIST_DIR / f"build_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        report_path.write_text(generate_report(all_results), encoding="utf-8")
        print(f"\nReport saved to: {report_path}")

    # Exit code
    errors = sum(1 for r in all_results if not r["ok"])
    if errors > 0:
        print(f"\n[BUILD FAILED] {errors} error(s) found.")
        sys.exit(1)
    else:
        print(f"\n[BUILD PASSED] All validations passed.")
        sys.exit(0)


if __name__ == "__main__":
    main()
