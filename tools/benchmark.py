#!/usr/bin/env python3
"""
Speedy Scandocs — naming-accuracy benchmark harness (DEVELOPER TOOL ONLY)
===========================================================================

This script is never bundled into the app. It exists so a developer can
answer one question: "did this change to the prompt / config / matching
logic actually make document naming more accurate?"

WHAT IT DOES
------------
It runs the real classification pipeline (ConfigManager, ClientListManager,
DocumentExtractor, APIClient, FileProcessor) against a folder of documents
whose CURRENT FILENAMES ARE TRUSTED AS GROUND TRUTH — i.e. files an
employee has already reviewed and correctly named in the normal
"LAST, First - Description.ext" convention, e.g.:

    CONTRERAS, Francisco - Reduction Request to First Care.pdf

For each file it copies the document to a scratch temp directory (the
original is NEVER touched — see "SAFETY" below), feeds the copy through
the pipeline, and compares what the pipeline predicts against the
client/description parsed from the original filename.

BUILDING A GOOD BENCHMARK SET
------------------------------
The benchmark set must be STRATIFIED, not just a folder of easy, clean
scans. Include a proportional mix of the document conditions the tool
actually sees in production, e.g.:

    - phone photos (skewed, glare, partial pages)
    - incoming faxes (headers/footers, low contrast, cover sheets)
    - certified mail / USPS green-card scans
    - multi-page progress reports / IME reports
    - clean flatbed scans (the easy case)

A benchmark set made only of clean scans will make every change look
like an improvement. Aim for at least 20-30 files per category if you
can, and keep the mix roughly proportional to what actually crosses the
scanner in a normal week.

SAFETY — READ THIS
-------------------
  * The --input-dir you point this at contains REAL CLIENT DOCUMENTS.
  * This script NEVER modifies, renames, or deletes anything in
    --input-dir. Every file is copied to a temp directory before the
    pipeline (which normally renames files in place) ever touches it.
  * The input directory itself, and any CSV this script writes
    (--out / --compare files), MUST NEVER BE COMMITTED TO GIT. They
    contain real client names and case descriptions. This repo's
    .gitignore already excludes tools/benchmark_data/, tools/*.csv, and
    benchmark_results*.csv — keep your benchmark data under one of
    those paths, or add your own path to .gitignore before using it.

USAGE
-----
    # Basic run, prints a summary table to stdout
    python3 tools/benchmark.py --input-dir tools/benchmark_data/set1

    # Save per-file results for later comparison
    python3 tools/benchmark.py --input-dir tools/benchmark_data/set1 \\
        --out tools/baseline_results.csv

    # After making a change to the prompt/config, compare against baseline
    python3 tools/benchmark.py --input-dir tools/benchmark_data/set1 \\
        --out tools/after_results.csv --compare tools/baseline_results.csv

    # Quick smoke test on a handful of files, with progress output
    python3 tools/benchmark.py --input-dir tools/benchmark_data/set1 \\
        --limit 10 --verbose

Each file hits a local LLM (via APIClient -> Ollama/OpenWebUI), so a run
over a large benchmark set can be slow. Use --limit while iterating and
run the full set for a final before/after comparison.

EXIT CODE
---------
Always 0. This is a reporting tool, not a CI gate — a "worse" run is
information for the developer, not a build failure.
"""

import argparse
import copy
import csv
import difflib
import os
import shutil
import sys
import tempfile
import traceback
from dataclasses import dataclass
from typing import Optional

# ── Make scandocs_tool importable ──────────────────────────────────────────
# This script lives in tools/, scandocs_tool.py lives in the repo root, one
# directory up. Add that directory to sys.path so `import scandocs_tool`
# works regardless of the caller's current working directory.
_TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_TOOLS_DIR)
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

st = None  # populated on demand by _load_scandocs()


def _load_scandocs():
    """Import scandocs_tool lazily.

    Deferred so --help and argument validation work on machines that lack the
    app's GUI dependencies (tkinter/ttkbootstrap). Exits with a clear message
    rather than a bare traceback if the import fails.
    """
    global st
    if st is None:
        try:
            import scandocs_tool as _st
        except ImportError as exc:
            sys.exit(
                f"Could not import scandocs_tool: {exc}\n\n"
                "The benchmark runs the real pipeline, so it needs the app's\n"
                "dependencies (see requirements.txt). If tkinter is missing,\n"
                "run the benchmark on the same machine that runs the app."
            )
        st = _st
    return st


SUPPORTED_EXTS = {".pdf", ".jpg", ".jpeg"}


# ─────────────────────────────────────────────────────────────
# Ground truth parsing
# ─────────────────────────────────────────────────────────────

@dataclass
class GroundTruth:
    filename: str
    client: str
    description: str
    ext: str


def parse_ground_truth(filename: str) -> Optional[GroundTruth]:
    """Parse 'CLIENT - Description.ext' into components.

    Client is everything before the FIRST ' - '; description is
    everything after it, minus the extension. Returns None if the
    filename doesn't follow the convention (no ' - ' separator) — such
    files are skipped with a warning rather than crashing the run.
    """
    base, ext = os.path.splitext(filename)
    if " - " not in base:
        return None
    client, desc = base.split(" - ", 1)
    return GroundTruth(
        filename=filename,
        client=client.strip(),
        description=desc.strip(),
        ext=ext.lower(),
    )


# ─────────────────────────────────────────────────────────────
# Normalization / scoring helpers
# ─────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    """Case-insensitive, whitespace-normalized comparison key."""
    return " ".join((s or "").split()).casefold()


def exact_match(a: str, b: str) -> bool:
    return _norm(a) == _norm(b)


def fuzzy_ratio(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, _norm(a), _norm(b)).ratio()


# ─────────────────────────────────────────────────────────────
# Per-file result row
# ─────────────────────────────────────────────────────────────

CSV_FIELDS = [
    "filename",
    "expected_client",
    "predicted_client",
    "client_ok",
    "expected_desc",
    "predicted_desc",
    "desc_ok",
    "desc_similarity",
    "filename_ok",
    "extraction_method",
    "confidence",
    "doc_type",
    "status",
    "error",
]


def make_row(gt: GroundTruth, result: Optional["st.ProcessResult"],
             error: str = "") -> dict:
    predicted_client = getattr(result, "client", "") if result else ""
    predicted_desc = getattr(result, "description", "") if result else ""
    extraction_method = getattr(result, "extraction_method", "") if result else ""
    confidence = getattr(result, "confidence", "") if result else ""
    doc_type = getattr(result, "doc_type", "") if result else ""
    status = getattr(result, "status", "") if result else ""

    client_ok = bool(result) and exact_match(gt.client, predicted_client)
    desc_ok = bool(result) and exact_match(gt.description, predicted_desc)
    desc_similarity = fuzzy_ratio(gt.description, predicted_desc) if result else 0.0

    expected_full = f"{gt.client} - {gt.description}{gt.ext}"
    predicted_full = f"{predicted_client} - {predicted_desc}{gt.ext}" if result else ""
    filename_ok = bool(result) and exact_match(expected_full, predicted_full)

    return {
        "filename": gt.filename,
        "expected_client": gt.client,
        "predicted_client": predicted_client,
        "client_ok": client_ok,
        "expected_desc": gt.description,
        "predicted_desc": predicted_desc,
        "desc_ok": desc_ok,
        "desc_similarity": round(desc_similarity, 4),
        "filename_ok": filename_ok,
        "extraction_method": extraction_method,
        "confidence": confidence,
        "doc_type": doc_type,
        "status": status,
        "error": error,
    }


# ─────────────────────────────────────────────────────────────
# Running the real pipeline against a temp copy of one file
# ─────────────────────────────────────────────────────────────

def run_pipeline_on_copy(src_path: str, gt: GroundTruth, config: dict,
                          client_list: list, scratch_dir: str, index: int):
    """Copy `src_path` into `scratch_dir` under a NEUTRAL filename and run
    FileProcessor.process_file against the copy.

    Two things make this safe/valid as a blind test:

    1. SAFETY: process_file renames files on disk. We only ever operate
       on a throwaway copy in a temp directory — the original ground-truth
       file in --input-dir is never opened for writing, moved, or renamed.

    2. BLINDNESS: the ground-truth filename itself (e.g.
       "CONTRERAS, Francisco - Reduction Request to First Care.pdf")
       already looks like a correctly-processed file. If we ran the
       pipeline against a copy that kept that name,
       FileProcessor._already_processed() would recognize the pattern
       and short-circuit to status="skipped" without ever calling the
       model — silently making every result "correct" by construction.
       So the copy is given a neutral scanner-style name
       ("scan_0001.pdf") that carries no hint of the expected answer,
       matching how files actually arrive in production (scanner/fax
       software does not name files "Client - Description").

    A config copy is also used, with `skip_already_processed` forced off
    as a second, independent guard against the same short-circuit — this
    protects the benchmark even if a future change alters
    _already_processed()'s matching logic. If a `dry_run` flag exists in
    config["automation"] by the time this runs, it is also honored
    (checked defensively with .get()), but the temp-copy approach above
    is what actually guarantees the input file is never touched, and
    that guarantee does not depend on any config flag existing.
    """
    neutral_name = f"scan_{index:04d}{gt.ext}"
    copy_path = os.path.join(scratch_dir, neutral_name)
    shutil.copy2(src_path, copy_path)

    run_config = copy.deepcopy(config)
    run_config.setdefault("processing", {})["skip_already_processed"] = False
    # Defensive: honor a dry-run flag if one has landed in config by now,
    # without assuming it exists.
    if run_config.get("automation", {}).get("dry_run"):
        pass  # nothing extra to do — process_file would need to respect
              # this itself; the temp-copy isolation already protects us.

    result = st.FileProcessor.process_file(copy_path, run_config, client_list)
    return result


# ─────────────────────────────────────────────────────────────
# Scoring summary
# ─────────────────────────────────────────────────────────────

def summarize(rows: list) -> dict:
    total = len(rows)
    errored = sum(1 for r in rows if r["error"])
    scored = [r for r in rows if not r["error"]]
    n = len(scored)

    def pct(count):
        return (100.0 * count / n) if n else 0.0

    client_ok = sum(1 for r in scored if r["client_ok"])
    desc_ok = sum(1 for r in scored if r["desc_ok"])
    filename_ok = sum(1 for r in scored if r["filename_ok"])
    avg_desc_sim = (sum(r["desc_similarity"] for r in scored) / n) if n else 0.0

    by_method = {}
    for r in scored:
        method = r["extraction_method"] or "unknown"
        by_method.setdefault(method, {"n": 0, "client_ok": 0, "desc_ok": 0, "filename_ok": 0})
        m = by_method[method]
        m["n"] += 1
        m["client_ok"] += int(r["client_ok"])
        m["desc_ok"] += int(r["desc_ok"])
        m["filename_ok"] += int(r["filename_ok"])

    return {
        "total": total,
        "errored": errored,
        "scored": n,
        "client_ok": client_ok,
        "client_pct": pct(client_ok),
        "desc_ok": desc_ok,
        "desc_pct": pct(desc_ok),
        "avg_desc_similarity": avg_desc_sim,
        "filename_ok": filename_ok,
        "filename_pct": pct(filename_ok),
        "by_method": by_method,
    }


def print_summary(summary: dict):
    print()
    print("=" * 62)
    print("BENCHMARK SUMMARY")
    print("=" * 62)
    print(f"  Total files:          {summary['total']}")
    print(f"  Errored (excluded):   {summary['errored']}")
    print(f"  Scored:               {summary['scored']}")
    print("-" * 62)
    print(f"  Client exact match:   {summary['client_ok']}/{summary['scored']} "
          f"({summary['client_pct']:.1f}%)")
    print(f"  Description exact:    {summary['desc_ok']}/{summary['scored']} "
          f"({summary['desc_pct']:.1f}%)")
    print(f"  Description fuzzy avg:{summary['avg_desc_similarity']*100:6.1f}%")
    print(f"  Full filename exact:  {summary['filename_ok']}/{summary['scored']} "
          f"({summary['filename_pct']:.1f}%)")
    print("-" * 62)
    print("  By extraction method:")
    if not summary["by_method"]:
        print("    (no scored results)")
    for method, m in sorted(summary["by_method"].items()):
        n = m["n"]
        cp = (100.0 * m["client_ok"] / n) if n else 0.0
        dp = (100.0 * m["desc_ok"] / n) if n else 0.0
        fp = (100.0 * m["filename_ok"] / n) if n else 0.0
        print(f"    {method:10s} n={n:4d}  client={cp:5.1f}%  "
              f"desc={dp:5.1f}%  filename={fp:5.1f}%")
    print("=" * 62)
    print()


# ─────────────────────────────────────────────────────────────
# Baseline comparison / regression report
# ─────────────────────────────────────────────────────────────

def load_csv_rows(path: str) -> dict:
    """Load a previous run's CSV into {filename: row_dict}."""
    rows = {}
    with open(path, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows[row["filename"]] = row
    return rows


def _as_bool(v) -> bool:
    return str(v).strip().lower() in ("true", "1", "yes")


def print_regression_report(current_rows: list, baseline_path: str):
    baseline = load_csv_rows(baseline_path)
    current = {r["filename"]: r for r in current_rows}

    common = sorted(set(baseline) & set(current))
    only_in_current = sorted(set(current) - set(baseline))
    only_in_baseline = sorted(set(baseline) - set(current))

    print()
    print("#" * 62)
    print(f"REGRESSION REPORT  (baseline: {baseline_path})")
    print("#" * 62)

    if not common:
        print("  No filenames in common between current run and baseline — "
              "nothing to compare.")
        print("#" * 62)
        print()
        return

    fields_to_track = ["client_ok", "desc_ok", "filename_ok"]
    net = {f: {"improved": 0, "regressed": 0, "unchanged": 0} for f in fields_to_track}
    improved_files = {f: [] for f in fields_to_track}
    regressed_files = {f: [] for f in fields_to_track}

    for fname in common:
        b = baseline[fname]
        c = current[fname]
        for field in fields_to_track:
            b_ok = _as_bool(b.get(field, ""))
            c_ok = _as_bool(c.get(field, ""))
            if b_ok == c_ok:
                net[field]["unchanged"] += 1
            elif c_ok and not b_ok:
                net[field]["improved"] += 1
                improved_files[field].append(fname)
            else:
                net[field]["regressed"] += 1
                regressed_files[field].append(fname)

    print(f"  Files compared: {len(common)}")
    if only_in_current:
        print(f"  New files (not in baseline, excluded from diff): {len(only_in_current)}")
    if only_in_baseline:
        print(f"  Files missing vs. baseline (excluded from diff): {len(only_in_baseline)}")
    print("-" * 62)

    for field in fields_to_track:
        stats = net[field]
        delta = stats["improved"] - stats["regressed"]
        sign = "+" if delta >= 0 else ""
        print(f"  {field:14s} improved={stats['improved']:4d}  "
              f"regressed={stats['regressed']:4d}  "
              f"unchanged={stats['unchanged']:4d}   net={sign}{delta}")

    # Description similarity: numeric delta, not just boolean
    sims_delta = []
    for fname in common:
        try:
            b_sim = float(baseline[fname].get("desc_similarity", 0) or 0)
            c_sim = float(current[fname].get("desc_similarity", 0) or 0)
            sims_delta.append(c_sim - b_sim)
        except ValueError:
            continue
    if sims_delta:
        avg_delta = sum(sims_delta) / len(sims_delta)
        sign = "+" if avg_delta >= 0 else ""
        print(f"  {'desc_similarity':14s} avg change: {sign}{avg_delta*100:.2f} pts")

    print("-" * 62)
    print("  Regressed files (got worse):")
    any_regressions = False
    for field in fields_to_track:
        for fname in regressed_files[field]:
            any_regressions = True
            print(f"    [{field}] {fname}")
    if not any_regressions:
        print("    (none)")

    print("-" * 62)
    print("  Improved files (got better):")
    any_improvements = False
    for field in fields_to_track:
        for fname in improved_files[field]:
            any_improvements = True
            print(f"    [{field}] {fname}")
    if not any_improvements:
        print("    (none)")

    print("#" * 62)
    print()


# ─────────────────────────────────────────────────────────────
# CSV writing
# ─────────────────────────────────────────────────────────────

def write_csv(rows: list, out_path: str):
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


# ─────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────

def print_startup_warning(input_dir: str, out_path: Optional[str]):
    print("=" * 62)
    print("  WARNING — this tool reads REAL CLIENT DOCUMENTS")
    print("=" * 62)
    print(f"  Input directory: {input_dir}")
    if out_path:
        print(f"  Output CSV:      {out_path}")
    print()
    print("  Never commit the input directory or any CSV output from this")
    print("  script to git — they contain real client names and case")
    print("  descriptions. Keep benchmark data under tools/benchmark_data/")
    print("  or another path already excluded in .gitignore.")
    print("=" * 62)
    print()


def collect_input_files(input_dir: str, limit: Optional[int]) -> list:
    files = []
    for name in sorted(os.listdir(input_dir)):
        path = os.path.join(input_dir, name)
        if not os.path.isfile(path):
            continue
        if os.path.splitext(name)[1].lower() not in SUPPORTED_EXTS:
            continue
        files.append(path)
    if limit is not None:
        files = files[:limit]
    return files


def main():
    parser = argparse.ArgumentParser(
        description="Benchmark Speedy Scandocs naming accuracy against a "
                     "folder of already-correctly-named real documents.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--input-dir", required=True,
                         help="Directory of correctly-named ground-truth documents "
                              "(PDF/JPG). Never modified.")
    parser.add_argument("--out", default=None,
                         help="Write per-file results to this CSV path.")
    parser.add_argument("--compare", default=None,
                         help="Path to a previous run's CSV; prints a regression "
                              "report against it.")
    parser.add_argument("--limit", type=int, default=None,
                         help="Only process the first N files (alphabetical).")
    parser.add_argument("--verbose", action="store_true",
                         help="Print progress per file (slow — each file hits a "
                              "local LLM).")
    args = parser.parse_args()

    input_dir = os.path.abspath(args.input_dir)
    if not os.path.isdir(input_dir):
        print(f"ERROR: --input-dir does not exist or is not a directory: {input_dir}")
        sys.exit(0)

    print_startup_warning(input_dir, args.out)

    files = collect_input_files(input_dir, args.limit)
    if not files:
        print(f"No PDF/JPG files found in {input_dir}. Nothing to do.")
        sys.exit(0)

    # Import the app now that we know we have real work to do.
    _load_scandocs()

    # Load real config + client list, exactly as the app would.
    config_manager = st.ConfigManager()
    config = config_manager.config
    client_list_path = config.get("paths", {}).get("client_list_file", "")
    client_list = st.ClientListManager.load(client_list_path)
    if not client_list:
        print(f"WARNING: client list at '{client_list_path}' is empty or missing. "
              "Client matching will fail for every file until this is fixed "
              "(check config.json's paths.client_list_file).")

    rows = []
    with tempfile.TemporaryDirectory(prefix="scandocs_benchmark_") as scratch_dir:
        for i, src_path in enumerate(files, start=1):
            filename = os.path.basename(src_path)
            gt = parse_ground_truth(filename)
            if gt is None:
                print(f"SKIP (doesn't match 'Client - Description.ext'): {filename}")
                continue

            if args.verbose:
                print(f"[{i}/{len(files)}] {filename} ...", flush=True)

            error = ""
            result = None
            try:
                result = run_pipeline_on_copy(
                    src_path, gt, config, client_list, scratch_dir, i
                )
                if getattr(result, "status", "") == "error":
                    error = getattr(result, "error_message", "") or "unknown pipeline error"
            except Exception as e:
                error = f"{type(e).__name__}: {e}"
                if args.verbose:
                    traceback.print_exc()

            row = make_row(gt, result, error=error)
            rows.append(row)

            if args.verbose:
                status = "ERROR" if error else (
                    "OK" if row["filename_ok"] else "MISMATCH"
                )
                print(f"    -> {status}  client_ok={row['client_ok']} "
                      f"desc_sim={row['desc_similarity']:.2f}")

    summary = summarize(rows)
    print_summary(summary)

    if args.out:
        out_path = os.path.abspath(args.out)
        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
        write_csv(rows, out_path)
        print(f"Wrote {len(rows)} rows to {out_path}")

    if args.compare:
        print_regression_report(rows, args.compare)

    sys.exit(0)


if __name__ == "__main__":
    main()
