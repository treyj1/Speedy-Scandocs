#!/usr/bin/env python3
"""
Speedy Scandocs
Automatically classifies and renames scanned law firm documents
using a local AI model (OpenWebUI / Ollama).

Usage:
    python scandocs_tool.py

Requirements:
    pip install -r requirements.txt
"""

import os
import re
import sys
import json
import base64
import difflib
import hashlib
import logging
import threading
import queue
import csv
import datetime
import subprocess

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    _XLSX_AVAILABLE = True
except ImportError:
    _XLSX_AVAILABLE = False
from dataclasses import dataclass, field
from typing import Optional, List

import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import requests
except ImportError:
    requests = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

try:
    from PIL import Image as PILImage, ImageTk as PILImageTk
except ImportError:
    PILImage = None
    PILImageTk = None


# ─────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(SCRIPT_DIR, "assets")

# ── Version + auto-update ──────────────────────────────────────────────────
# APP_VERSION is bumped by build/release.py — keep it in sync with the
# installer.iss AppVersion. Auto-update checks GitHub Releases on UPDATE_REPO
# and compares the latest tag (vX.Y.Z) against APP_VERSION.
APP_VERSION = "1.9.1"
UPDATE_REPO = "treyj1/Speedy-Scandocs"
UPDATE_API_URL = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
UPDATE_CHECK_INTERVAL_SEC = 24 * 60 * 60   # 24 hours

# ── Build flavor ────────────────────────────────────────────────────────────
# "production" or "test". A test build uses a separate app-data directory
# (so it never touches a real installation's config/logs/reports/client
# list) and disables auto-update entirely. Do not change this by editing the
# installer/spec files in this pass — it's a source-level switch for now.
BUILD_FLAVOR = "production"   # "production" or "test"
IS_TEST_BUILD = (BUILD_FLAVOR == "test")
APP_TITLE = "Speedy Scandocs" + (" [TEST BUILD]" if IS_TEST_BUILD else "")
_APP_DATA_DIRNAME = "SpeedyScandocs-Test" if IS_TEST_BUILD else "SpeedyScandocs"

# ── Colour-palette definitions ─────────────────────────────────────────────
# App primary color — fixed, no user-selectable palette
_APP_PRIMARY   = "#1565c0"
_APP_LIGHT     = "#e3f2fd"
_APP_MID       = "#1976d2"

# ── Typography ─────────────────────────────────────────────────────────────
# Single app-wide font family.
APP_FONT = "Times New Roman"

# ── User-writable data directory ───────────────────────────────────────────
# When installed to Program Files / Applications the app bundle is read-only,
# so config, logs, and reports are stored in the platform user-data folder.
if getattr(sys, "frozen", False):
    if sys.platform == "win32":
        _appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
    elif sys.platform == "darwin":
        _appdata = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        _appdata = os.environ.get("XDG_DATA_HOME", os.path.join(os.path.expanduser("~"), ".local", "share"))
    _USER_DATA_DIR = os.path.join(_appdata, _APP_DATA_DIRNAME)
    os.makedirs(_USER_DATA_DIR, exist_ok=True)
else:
    _USER_DATA_DIR = SCRIPT_DIR

CONFIG_PATH            = os.path.join(_USER_DATA_DIR, "config.json")
LOG_PATH               = os.path.join(_USER_DATA_DIR, "scandocs_tool.log")
DEFAULT_REPORTS_FOLDER = os.path.join(_USER_DATA_DIR, "Reports")


def _open_file(path: str):
    """Open a file with the default system viewer — cross-platform."""
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


# ── PyInstaller bundle: point pytesseract at the bundled Tesseract binary ──
_bundle_dir = getattr(sys, "_MEIPASS", None)
if _bundle_dir:
    try:
        import pytesseract as _pt
        # Windows bundle: Tesseract-OCR/tesseract.exe
        _bundled_tess_win = os.path.join(_bundle_dir, "Tesseract-OCR", "tesseract.exe")
        # Mac bundle: tesseract (no extension)
        _bundled_tess_mac = os.path.join(_bundle_dir, "tesseract")
        if os.path.isfile(_bundled_tess_win):
            _pt.pytesseract.tesseract_cmd = _bundled_tess_win
            os.environ["TESSDATA_PREFIX"] = os.path.join(_bundle_dir, "Tesseract-OCR", "tessdata")
        elif os.path.isfile(_bundled_tess_mac):
            _pt.pytesseract.tesseract_cmd = _bundled_tess_mac
            os.environ["TESSDATA_PREFIX"] = os.path.join(_bundle_dir, "tessdata")
    except ImportError:
        pass
else:
    # Not bundled — try known system install locations (Windows + Mac Homebrew)
    _TESS_CANDIDATES = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        "/opt/homebrew/bin/tesseract",   # Mac Apple Silicon (Homebrew)
        "/usr/local/bin/tesseract",      # Mac Intel (Homebrew)
    ]
    try:
        import pytesseract as _pt
        for _p in _TESS_CANDIDATES:
            if os.path.isfile(_p):
                _pt.pytesseract.tesseract_cmd = _p
                break
    except ImportError:
        pass


# ── Windows taskbar identity ───────────────────────────────────────────────
# Without an explicit AppUserModelID, Windows groups the running app under
# Python's default ID and the taskbar icon falls back to Python's, not the
# embedded EXE icon. Must be set before any Tk window is mapped.
if sys.platform == "win32":
    try:
        import ctypes as _ctypes
        _ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "com.gdj.speedyscandocs"
        )
    except Exception:
        pass


# ── Bundled font registration ──────────────────────────────────────────────
# Must run before any Tk window is created — GDI/CoreText caches font lists
# at Tk init time, so a late registration won't be visible to tkinter.
def _load_bundled_fonts() -> None:
    fonts_dir = os.path.join(ASSETS_DIR, "fonts")
    if not os.path.isdir(fonts_dir):
        return
    ttfs = [os.path.join(fonts_dir, f) for f in os.listdir(fonts_dir)
            if f.lower().endswith(".ttf")]
    if not ttfs:
        return
    if sys.platform == "win32":
        try:
            import ctypes
            FR_PRIVATE = 0x10
            gdi32 = ctypes.WinDLL("gdi32")
            gdi32.AddFontResourceExW.argtypes = [
                ctypes.c_wchar_p, ctypes.c_ulong, ctypes.c_void_p]
            gdi32.AddFontResourceExW.restype = ctypes.c_int
            for path in ttfs:
                if gdi32.AddFontResourceExW(path, FR_PRIVATE, 0) == 0:
                    logging.info(f"AddFontResourceExW failed: {path}")
        except Exception as e:
            logging.info(f"Windows font load failed: {e}")
    elif sys.platform == "darwin":
        try:
            import ctypes
            from ctypes import c_void_p, c_bool, c_long, c_uint32, c_char_p
            cf = ctypes.CDLL("/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation")
            ct = ctypes.CDLL("/System/Library/Frameworks/CoreText.framework/CoreText")
            cf.CFStringCreateWithCString.restype = c_void_p
            cf.CFStringCreateWithCString.argtypes = [c_void_p, c_char_p, c_uint32]
            cf.CFURLCreateWithFileSystemPath.restype = c_void_p
            cf.CFURLCreateWithFileSystemPath.argtypes = [c_void_p, c_void_p, c_long, c_bool]
            cf.CFRelease.argtypes = [c_void_p]
            ct.CTFontManagerRegisterFontsForURL.restype = c_bool
            ct.CTFontManagerRegisterFontsForURL.argtypes = [c_void_p, c_uint32, c_void_p]
            kCFStringEncodingUTF8 = 0x08000100
            kCTFontManagerScopeProcess = 1
            for path in ttfs:
                cfstr = cf.CFStringCreateWithCString(
                    None, path.encode("utf-8"), kCFStringEncodingUTF8)
                if not cfstr:
                    continue
                cfurl = cf.CFURLCreateWithFileSystemPath(None, cfstr, 0, False)
                if cfurl:
                    ct.CTFontManagerRegisterFontsForURL(
                        cfurl, kCTFontManagerScopeProcess, None)
                    cf.CFRelease(cfurl)
                cf.CFRelease(cfstr)
        except Exception as e:
            logging.info(f"macOS font load failed: {e}")

DEFAULT_CONFIG: dict = {
    "paths": {
        "scandocs_folder": "",
        "client_list_file": os.path.join(_USER_DATA_DIR, "client_list.txt"),
    },
    "api": {
        "openwebui_url": "http://localhost:3000",
        "ollama_url": "http://localhost:11434",
        "model": "llama3.2-vision",
        "api_key": "",
        "timeout_connect": 10,
        "timeout_read": 120,
    },
    "processing": {
        "fuzzy_threshold": 0.82,
        "max_ocr_chars": 8000,
        "require_high_confidence": True,
        "max_pages": 5,
        "skip_already_processed": True,
        "audit_mode": True,
        "file_mode": False,
        "file_mode_destination": "",
        "suggest_location_enabled": False,
        "suggest_location_parent_folder": "",
        "auto_commit_moves": False,
        "candidate_list_size": 10,
        "show_manual_entry_tab": False,  # legacy standalone Manual Entry tab
        "extraction_method": "ocr",   # "ocr" or "vision"
        "max_vision_pages": 2,          # pages sent to vision model per doc
        "ocr_preprocess": True,         # upscale/binarize/autocontrast before OCR
    },
    "reports": {
        "auto_save": True,
        "report_folder": DEFAULT_REPORTS_FOLDER,
    },
    "updates": {
        "check_on_startup": True,
        "last_check_iso": "",
        "skip_version": "",
    },
    "ui": {
        "preview_popup_width": 0,   # 0 = not yet set by the user; use the default large size
        "preview_popup_height": 0,
    },
    # ── Foundation for later passes — all defaults below reproduce today's
    # behavior exactly. No UI wired to these yet.
    "naming": {
        "preserve_acronyms": True,          # safe, pure improvement — on by default
        "include_recipient": False,
        "include_doc_date": False,
        "templates": {},                    # doc_type -> template string, filled later
        "default_template": "{client} - {doc_type}",
        "date_disambiguation": False,       # False = legacy "(1)" behavior
        "unknown_client_label": "A-UNKNOWN CLIENT",
        "no_client_label": "A-NEEDS REVIEW",   # keep legacy default
    },
    "reading": {
        "skip_fax_cover_pages": False,
        "deskew_photos": False,
        "vision_escalation": False,
        "extract_claim_numbers": False,
    },
    "learning": {
        "log_corrections": True,            # foundation; harmless, just writes a log
        "document_types": "off",            # "off" | "suggest" | "auto"
        "client_relationships": "off",      # "off" | "suggest" | "auto"
        "claim_linking": "off",             # "off" | "suggest" | "auto"
        "observations_required": 3,
        "retroactive_rename": "off",        # "off" | "preview"
    },
    "automation": {
        "dry_run": False,
        "watch_folder": False,
        "watch_poll_seconds": 15,
    },
    "safety": {
        # Deliberately NOT user-toggleable. Present for diagnostics/logging only.
        "instance_lock": True,
        "undo_log": True,
        "recheck_before_rename": True,
    },
}

SUPPORTED_EXTENSIONS = {".pdf", ".jpg", ".jpeg"}
ILLEGAL_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# ── Vision model allowlist ───────────────────────────────────────────
# Substring-matched against the lowercased Ollama model name. Any model
# whose tag contains one of these prefixes is treated as vision-capable,
# which unlocks "Use Vision Model" in Settings. Excluded tags (cloud
# variants) send document images off-device, which isn't acceptable for
# law office client documents, so we force those to OCR mode.
VISION_MODEL_PREFIXES = [
    "llama3.2-vision",  # 11b and 90b
    "gemma4",           # latest, e2b, e4b, 26b, 31b (local variants)
]
VISION_MODEL_EXCLUSIONS = ["-cloud", ":cloud"]


def _disable_combobox_scroll(widget) -> None:
    """Stop a ttk.Combobox from cycling values when the mouse wheel scrolls
    over it. Otherwise an accidental scroll while hovering Settings can flip
    Use OCR ↔ Use Vision Model without the user realizing."""
    for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        widget.bind(seq, lambda _e: "break")


def model_supports_vision(model_name: str) -> bool:
    """Return True if the given Ollama model tag is a local vision model
    in our allowlist. Cloud variants are intentionally excluded."""
    if not model_name:
        return False
    name = model_name.lower()
    if any(excl in name for excl in VISION_MODEL_EXCLUSIONS):
        return False
    return any(prefix in name for prefix in VISION_MODEL_PREFIXES)


# ─────────────────────────────────────────────────────────────
# Data classes
# ─────────────────────────────────────────────────────────────

@dataclass
class ExtractionResult:
    content_type: str   # "text" or "image"
    content: str        # text string or first base64 image (back-compat)
    mime_type: str = "image/png"
    method: str = ""    # "pymupdf", "tesseract", or "vision"
    # For multi-page vision extraction. When non-empty, APIClient sends
    # every base64 image in this list to the model. `content` mirrors
    # images[0] for back-compat with single-image code paths.
    images: List[str] = field(default_factory=list)


@dataclass
class ProcessResult:
    original_name: str
    final_name: str
    status: str         # "renamed", "needs_review", "skipped", "error"
    client: str = ""
    description: str = ""
    confidence: str = ""
    error_message: Optional[str] = None
    renamed_at: Optional[str] = None
    extraction_method: str = ""
    # Audit fields — filled in by employee after processing
    audit_correct: bool = False
    audit_wrong_client: bool = False
    audit_bad_description: bool = False
    audit_failed_client: bool = False
    audit_should_review: bool = False
    audit_corrected_name: str = ""  # what the employee said it should be named
    # File mode
    moved_to: str = ""              # destination path after a successful move
    pending_dest: str = ""          # destination staged via "Apply to Selected", not yet moved
    # Foundation fields for later passes (learning, structured naming, etc.)
    raw_client: str = ""            # what the model returned, before fuzzy matching
    raw_confidence: str = ""
    extracted_text: str = ""        # first 2000 chars of what was actually read
    doc_hash: str = ""              # sha256 of file bytes, for dedupe
    doc_type: str = ""              # structured field, populated in a later pass
    recipient: str = ""
    doc_date: str = ""
    direction: str = ""             # "incoming" | "outgoing" | ""
    claim_number: str = ""
    was_dry_run: bool = False
    skip_reason: str = ""           # why a file was skipped, for reporting


# ─────────────────────────────────────────────────────────────
# ConfigManager
# ─────────────────────────────────────────────────────────────

class ConfigManager:
    def __init__(self):
        self.config = self._load()

    def _load(self) -> dict:
        if not os.path.exists(CONFIG_PATH):
            self._write(DEFAULT_CONFIG)
            return self._deep_copy(DEFAULT_CONFIG)
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return self._deep_merge(self._deep_copy(DEFAULT_CONFIG), data)
        except Exception as e:
            logging.warning(f"Could not load config.json: {e}. Using defaults.")
            return self._deep_copy(DEFAULT_CONFIG)

    @staticmethod
    def _deep_merge(defaults: dict, saved: dict) -> dict:
        """Recursively merge `saved` (loaded from disk) over `defaults`.

        - A saved scalar overrides the default.
        - A key present in defaults but missing from saved keeps the default.
        - A key present in saved but unknown to defaults is preserved as-is
          (never silently dropped — e.g. a section added by a newer version
          of the app, or user data from a future config format).
        - Nested dicts are merged recursively rather than replaced wholesale.

        This replaces a previous implementation that merged section-by-section
        with hardcoded lines (`merged["paths"].update(...)`, etc.) — any
        section added to DEFAULT_CONFIG without a matching hardcoded line
        there would have its saved values silently discarded on every load.
        """
        if not isinstance(saved, dict):
            return defaults
        merged = dict(defaults)
        for key, saved_val in saved.items():
            if key in merged and isinstance(merged[key], dict) and isinstance(saved_val, dict):
                merged[key] = ConfigManager._deep_merge(merged[key], saved_val)
            else:
                merged[key] = saved_val
        return merged

    def save(self, data: dict = None):
        if data is None:
            data = self.config
        self._write(data)
        self.config = data

    def _write(self, data: dict):
        tmp = CONFIG_PATH + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            os.replace(tmp, CONFIG_PATH)
        except Exception as e:
            logging.error(f"Failed to write config: {e}")
            raise

    def validate(self) -> list:
        errors = []
        scandocs = self.config["paths"].get("scandocs_folder", "")
        if not scandocs:
            errors.append("Scandocs folder is not configured.")
        elif not os.path.isdir(scandocs):
            errors.append(f"Scandocs folder not found:\n  {scandocs}")
        client_file = self.config["paths"].get("client_list_file", "")
        if not client_file:
            errors.append("Client list file path is not configured.")
        return errors

    @staticmethod
    def _deep_copy(d: dict) -> dict:
        return json.loads(json.dumps(d))


# ─────────────────────────────────────────────────────────────
# ClientListManager
# ─────────────────────────────────────────────────────────────

class ClientListManager:

    @staticmethod
    def load(path: str) -> list:
        """Load client list and automatically add space-variant for any hyphenated names.
        e.g. 'GARCIA-TELLEZ, Miguel' also registers as 'GARCIA TELLEZ, Miguel' internally
        so OCR output (which often drops hyphens) still fuzzy-matches correctly."""
        if not path or not os.path.isfile(path):
            return []
        try:
            with open(path, "r", encoding="utf-8") as f:
                names = [
                    line.strip()
                    for line in f
                    if line.strip() and not line.startswith("#")
                ]
            # Add space-variants for hyphenated names (kept in order, deduped)
            result = []
            seen = set()
            for name in names:
                if name not in seen:
                    result.append(name)
                    seen.add(name)
                if '-' in name:
                    variant = name.replace('-', ' ')
                    if variant not in seen:
                        result.append(variant)
                        seen.add(variant)
            return result
        except Exception as e:
            logging.error(f"Could not load client list from {path}: {e}")
            return []

    @staticmethod
    def save(path: str, clients: list):
        with open(path, "w", encoding="utf-8") as f:
            for name in sorted(set(clients)):
                f.write(name + "\n")

    @staticmethod
    def _invert_name(name: str) -> str:
        """Try to convert 'First Last' → 'LAST, First' for matching.
        If the name already contains a comma, returns it unchanged."""
        if ',' in name:
            return name
        parts = name.strip().split()
        if len(parts) >= 2:
            last = parts[-1].upper()
            first = ' '.join(parts[:-1])
            return f"{last}, {first}"
        return name

    @staticmethod
    def _normalize(name: str) -> str:
        """Lowercase and replace hyphens with spaces for comparison.
        Allows 'Garcia-Tellez' to match 'GARCIA TELLEZ' and vice versa."""
        return name.lower().replace('-', ' ')

    @staticmethod
    def _all_inversions(name: str) -> list:
        """Generate all possible LAST, First splits for a multi-word name.
        e.g. "Julio Solano Trujillo" →
             ["TRUJILLO, Julio Solano",   ← standard (last word = last name)
              "SOLANO TRUJILLO, Julio"]   ← two-word last name
        Only adds extra splits for 3+ word names; single-last-word is always first.
        Excludes splits where first-name portion would be empty."""
        if ',' in name:
            return [name]   # already in LAST, First form
        parts = name.strip().split()
        if len(parts) < 2:
            return [name]
        results = []
        # Try each split point from right to left (most common last-name first)
        for split in range(len(parts) - 1, 0, -1):
            last  = ' '.join(parts[split:]).upper()
            first = ' '.join(parts[:split])
            results.append(f"{last}, {first}")
        return results

    @staticmethod
    def fuzzy_match(candidate: str, client_list: list, threshold: float = 0.82) -> Optional[str]:
        if not candidate or not client_list:
            return None
        candidate = candidate.strip().strip(".,;:'\"").strip()

        # Build normalized index (lowercase + hyphens→spaces)
        norm_list = [ClientListManager._normalize(e) for e in client_list]

        # 1. Exact match after normalization — always preferred
        candidate_norm = ClientListManager._normalize(candidate)
        for i, entry_norm in enumerate(norm_list):
            if entry_norm == candidate_norm:
                return client_list[i]

        # 2. Try all inversion forms — standard split first, then multi-word last names.
        inversions = ClientListManager._all_inversions(candidate)

        # Also generate truncated-last-name forms:
        # "SOLANO TRUJILLO, Julio" → also try "SOLANO, Julio"
        # This handles AI returning 3-word names where only one surname is in the list.
        truncated = set()
        for inv in inversions:
            if ',' in inv:
                last_part, first_part = inv.split(',', 1)
                last_words = last_part.strip().split()
                if len(last_words) > 1:
                    truncated.add(f"{last_words[0]}, {first_part.strip()}")

        for inv in list(inversions) + list(truncated):
            inv_norm = ClientListManager._normalize(inv)
            if inv_norm == candidate_norm:
                continue
            for i, entry_norm in enumerate(norm_list):
                if entry_norm == inv_norm:
                    return client_list[i]   # exact match on this inversion

        # 3. Fuzzy match — score all inversion forms including truncated, keep best
        all_forms = {candidate_norm} | {
            ClientListManager._normalize(inv) for inv in list(inversions) + list(truncated)
        }
        best_entry = None
        best_score = 0.0
        for form in all_forms:
            for i, entry_norm in enumerate(norm_list):
                score = difflib.SequenceMatcher(None, form, entry_norm).ratio()
                if score > best_score:
                    best_score = score
                    best_entry = client_list[i]

        if best_entry and best_score >= threshold:
            return best_entry

        # 4. Compound surname prefix match — handles OCR merging like
        #    "Delgadillocuellar" where OCR drops the space between two last names.
        #    Extract the last-name portion of the candidate (before comma, or last word).
        if ',' in candidate:
            cand_last = candidate.split(',')[0].strip()
        else:
            parts = candidate.strip().split()
            cand_last = parts[-1] if parts else candidate
        cand_last_norm = ClientListManager._normalize(cand_last)

        # cand_last must be strictly longer than the entry's last name — otherwise
        # this is just a plain last-name match, not a merged compound surname.
        if len(cand_last_norm) >= 7:
            # Extract candidate's first name for cross-checking
            if ',' in candidate:
                cand_first_norm = ClientListManager._normalize(candidate.split(',')[1].strip())
            else:
                parts = candidate.strip().split()
                cand_first_norm = ClientListManager._normalize(parts[0]) if len(parts) > 1 else ""

            prefix_matches = []
            for i, entry in enumerate(client_list):
                entry_last = ClientListManager._normalize(
                    entry.split(',')[0].strip() if ',' in entry else entry.split()[-1]
                )
                # Require cand_last is longer (genuinely merged surname, not same last name)
                if len(entry_last) >= 6 and len(cand_last_norm) > len(entry_last) and cand_last_norm.startswith(entry_last):
                    # Cross-check first name if we have one
                    if cand_first_norm:
                        entry_first_norm = ClientListManager._normalize(
                            entry.split(',')[1].strip() if ',' in entry else ""
                        )
                        # First names must share a common start (handles abbreviations/nicknames)
                        if entry_first_norm and not (
                            entry_first_norm.startswith(cand_first_norm) or
                            cand_first_norm.startswith(entry_first_norm)
                        ):
                            continue  # First name mismatch — skip
                    prefix_matches.append(i)

            # Only act when exactly one client matches — ambiguity → NEEDS_REVIEW
            if len(prefix_matches) == 1:
                logging.info(
                    f"fuzzy_match: prefix match '{cand_last_norm}' → "
                    f"'{client_list[prefix_matches[0]]}'"
                )
                return client_list[prefix_matches[0]]

        return None

    @staticmethod
    def is_valid_format(name: str) -> bool:
        """Accepts 'LAST, First' or 'LAST, First Middle' etc."""
        return bool(re.match(r"^[A-Za-z\-\'\. ]+,\s+[A-Za-z\-\'\. ]+$", name.strip()))

    @staticmethod
    def filter_candidates(doc_text: str, client_list: list, top_n: int = 10) -> list:
        """Pre-filter the client list to the top_n most likely candidates.

        Strategy: only extract text from labeled keyword regions (Client:, RE:,
        Insured:, etc.) — NOT the full document.  Using the full document caused
        the original anchoring failure (facility name 'Velazquez Pain Summerlin'
        → model picked VELAZQUEZ, Aida).  Restricting to labeled regions gives a
        clean signal of who the document is actually about.

        Falls back to the full list when no labeled regions are found.
        """
        if not doc_text or not client_list:
            return client_list

        # ── Step 1: extract text near labeled keywords only ───────────────────
        # Each pattern captures up to ~60 chars after the label.
        _label_re = re.compile(
            r"(?:Client(?:\s*/\s*Patient)?(?:\s+Name)?|Claimant|Clmt|Injured(?:\s+(?:Worker|Party))?|"
            r"Patient(?:\s+Name)?|Insured(?:\s+Name)?|Named\s+Insured|Employee|"
            r"RE|Re|Regarding|Subject)\s*[:\-\.]\s*"
            r"([A-Za-z][A-Za-z ,\-\'\.]{1,60})",
            re.IGNORECASE,
        )
        labeled_phrases = [m.group(1).strip() for m in _label_re.finditer(doc_text)]

        # ── Step 2: score every client against the labeled phrases ────────────
        if labeled_phrases:
            scored: list = []
            for client in client_list:
                client_lower = client.lower()
                last_name = client.split(",")[0].strip().lower()
                best = 0.0
                for phrase in labeled_phrases:
                    phrase_lower = phrase.lower()
                    # Exact last-name hit in the phrase → strong signal
                    if len(last_name) >= 3 and last_name in phrase_lower:
                        best = max(best, 0.90)
                        continue
                    ratio = difflib.SequenceMatcher(None, phrase_lower, client_lower).ratio()
                    if ratio > best:
                        best = ratio
                scored.append((best, client))

            scored.sort(key=lambda x: x[0], reverse=True)
            top_score = scored[0][0] if scored else 0.0

            if top_score >= 0.50:
                short_list = [name for _score, name in scored[:top_n]]
                logging.info(
                    f"filter_candidates: {len(client_list)} → {len(short_list)} candidates "
                    f"(top score {top_score:.2f}, phrases: {labeled_phrases})"
                )
                return short_list

        # No labeled regions found or no strong score → fall back to full list
        logging.info(
            f"filter_candidates: no labeled keyword regions found, using full list "
            f"({len(client_list)} clients)"
        )
        return client_list


# ─────────────────────────────────────────────────────────────
# DocumentExtractor
# ─────────────────────────────────────────────────────────────

class DocumentExtractor:
    """
    Extracts content from a document for AI classification.
    Returns an ExtractionResult with content_type="text" or "image".
    """

    IMAGE_RENDER_SCALE = 300 / 72  # ~4.17x — renders at 300 DPI for better OCR accuracy

    @staticmethod
    def extract(file_path: str, max_chars: int = 4000, max_pages: int = 5,
                vision_mode: bool = False, max_vision_pages: int = 2,
                ocr_preprocess: bool = True) -> ExtractionResult:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".pdf":
            if vision_mode:
                return DocumentExtractor._from_pdf_vision(file_path, max_vision_pages)
            return DocumentExtractor._from_pdf(file_path, max_chars, max_pages,
                                                ocr_preprocess=ocr_preprocess)
        elif ext in (".jpg", ".jpeg"):
            return DocumentExtractor._from_jpeg(file_path)
        else:
            raise ValueError(f"Unsupported file type: {ext}")

    @staticmethod
    def _from_pdf_vision(file_path: str, max_vision_pages: int) -> ExtractionResult:
        """Render the first N pages of a PDF as PNG images and return them
        as a base64 image list for the vision model. Bypasses OCR entirely."""
        if fitz is None:
            raise ImportError("PyMuPDF is not installed. Run: pip install PyMuPDF")
        doc = fitz.open(file_path)
        page_limit = min(doc.page_count, max(1, max_vision_pages))
        scale = DocumentExtractor.IMAGE_RENDER_SCALE
        mat = fitz.Matrix(scale, scale)
        images_b64: List[str] = []
        for i in range(page_limit):
            pix = doc[i].get_pixmap(matrix=mat)
            images_b64.append(base64.b64encode(pix.tobytes("png")).decode("utf-8"))
        doc.close()
        logging.info(
            f"{os.path.basename(file_path)}: vision mode — sending "
            f"{len(images_b64)} page(s) to model"
        )
        return ExtractionResult(
            content_type="image",
            content=images_b64[0] if images_b64 else "",
            mime_type="image/png",
            method="vision",
            images=images_b64,
        )

    # Labels that reliably identify the client when they appear in document text
    _CLIENT_LABEL_RE = re.compile(
        r"(?:Client(?:\s*/\s*Patient)?(?:\s+Name)?|Claimant|Clmt|"
        r"Injured(?:\s+(?:Worker|Party))?|Patient(?:\s+Name)?|"
        r"Insured(?:\s+Name)?|Named\s+Insured|Employee|"
        r"Your\s+Client|RE|Re|Regarding|Subject)\s*[:\-\.]\s*"
        r"([A-Za-z][A-Za-z ,\-\'\.]{1,60})",
        re.IGNORECASE,
    )

    # Words that signal the end of a name (stop words for phrase truncation)
    _NAME_STOP_RE = re.compile(
        r"\b(?:Law|Office|Attorney|Atty|DOB|Date|vs?\.?|and|LLC|Inc|"
        r"Insurance|Clinic|Hospital|Medical|Center|Corp|PC|LLP|PA)\b",
        re.IGNORECASE,
    )

    @staticmethod
    def _extract_labeled_snippets(page_text: str) -> list:
        """Extract name phrases following client-identifying labels on a page.
        Truncates each phrase at stop words so surrounding text doesn't bleed in."""
        snippets = []
        for m in DocumentExtractor._CLIENT_LABEL_RE.finditer(page_text):
            phrase = m.group(1).strip()
            # Truncate at stop words
            stop = DocumentExtractor._NAME_STOP_RE.search(phrase)
            if stop:
                phrase = phrase[:stop.start()].strip()
            phrase = phrase.strip(".,;: ")
            if len(phrase) >= 3:
                snippets.append(phrase)
        return snippets

    @staticmethod
    def _from_pdf(file_path: str, max_chars: int, max_pages: int = 5,
                  ocr_preprocess: bool = True) -> ExtractionResult:
        if fitz is None:
            raise ImportError("PyMuPDF is not installed. Run: pip install PyMuPDF")
        doc = fitz.open(file_path)
        page_limit = min(doc.page_count, max_pages)
        if doc.page_count > max_pages:
            logging.info(
                f"{os.path.basename(file_path)}: {doc.page_count} pages — "
                f"capping at {max_pages}"
            )

        # ── Pass 1: collect per-page text via PyMuPDF ─────────────────────
        page_texts = []
        all_native_text = ""
        for page in doc.pages(0, page_limit):
            pt = page.get_text("text").strip()
            page_texts.append(pt)
            all_native_text += pt

        has_native_text = len(all_native_text.strip()) >= 50

        if has_native_text:
            # ── Pass 2: find pages containing client-identifying labels ────
            labeled_pages = []
            unlabeled_pages = []
            for i, pt in enumerate(page_texts):
                snippets = DocumentExtractor._extract_labeled_snippets(pt)
                if snippets:
                    labeled_pages.append((i, pt, snippets))
                else:
                    unlabeled_pages.append((i, pt))

            if labeled_pages:
                # Build content: labeled pages first (with snippet annotation),
                # then remaining pages up to the char budget.
                parts = []
                for i, pt, snippets in labeled_pages:
                    parts.append(f"[Page {i+1} — labeled fields: {'; '.join(snippets)}]\n{pt}")
                for i, pt in unlabeled_pages:
                    parts.append(f"[Page {i+1}]\n{pt}")
                raw_text = "\n\n".join(parts).strip()[:max_chars]
                logging.info(
                    f"{os.path.basename(file_path)}: labeled pages found "
                    f"({[i+1 for i,_,_ in labeled_pages]}), sending those first"
                )
            else:
                # No labeled pages — send all pages sequentially
                raw_text = all_native_text.strip()[:max_chars]

            doc.close()
            return ExtractionResult(content_type="text", content=raw_text, method="pymupdf")

        # ── Image-only PDF: OCR every page up to limit, prioritize labeled ──
        doc.close()
        ocr_labeled = []
        ocr_unlabeled = []
        for i in range(page_limit):
            ocr_text = DocumentExtractor._ocr_pdf_page(file_path, page_index=i,
                                                        max_chars=max_chars,
                                                        preprocess=ocr_preprocess)
            if not ocr_text:
                continue
            snippets = DocumentExtractor._extract_labeled_snippets(ocr_text)
            if snippets:
                ocr_labeled.append((i, ocr_text, snippets))
            else:
                ocr_unlabeled.append((i, ocr_text))

        if ocr_labeled or ocr_unlabeled:
            parts = []
            for i, txt, snippets in ocr_labeled:
                parts.append(f"[Page {i+1} — labeled fields: {'; '.join(snippets)}]\n{txt}")
            for i, txt in ocr_unlabeled:
                parts.append(f"[Page {i+1}]\n{txt}")
            combined = "\n\n".join(parts).strip()[:max_chars]
            logging.info(
                f"OCR succeeded on {file_path} ({len(combined)} chars, "
                f"labeled pages: {[i+1 for i,_,_ in ocr_labeled]})"
            )
            return ExtractionResult(content_type="text", content=combined, method="tesseract")

        # Last resort: send page 1 as image to vision model
        return DocumentExtractor._render_pdf_page(file_path)

    @staticmethod
    def _ocr_pdf_page(file_path: str, page_index: int, max_chars: int,
                      preprocess: bool = True) -> str:
        """Run Tesseract OCR on a single PDF page rendered to an image.
        Returns extracted text if >= 50 characters were found, otherwise empty string.
        Returns empty string silently on any error so the caller can fall back gracefully."""
        if pytesseract is None or PILImage is None:
            return ""
        try:
            doc = fitz.open(file_path)
            page_index = min(page_index, doc.page_count - 1)
            pix = doc[page_index].get_pixmap(
                matrix=fitz.Matrix(DocumentExtractor.IMAGE_RENDER_SCALE,
                                   DocumentExtractor.IMAGE_RENDER_SCALE)
            )
            doc.close()
            import io
            pil = PILImage.open(io.BytesIO(pix.tobytes("png")))
            if preprocess:
                pil = DocumentExtractor._preprocess_for_ocr(pil)
            text = pytesseract.image_to_string(pil).strip()
            return text[:max_chars] if len(text) >= 50 else ""
        except Exception as e:
            logging.warning(f"OCR failed on {file_path} page {page_index}: {e}")
            return ""

    @staticmethod
    def _preprocess_for_ocr(img):
        """Upscale, contrast-normalize, and binarize a page image so Tesseract
        has the cleanest possible input. Falls back to the original image on
        any error — preprocessing must never break OCR."""
        try:
            from PIL import ImageOps
            g = img.convert("L")
            w, h = g.size
            # Tesseract is trained on ~300 DPI text; upscale small scans.
            target = 2000
            if max(w, h) < target:
                scale = target / max(w, h)
                g = g.resize((int(w * scale), int(h * scale)), PILImage.LANCZOS)
            g = ImageOps.autocontrast(g, cutoff=2)
            # Otsu threshold — picks the split that best separates ink from paper.
            hist = g.histogram()[:256]
            total = sum(hist)
            if total == 0:
                return g
            sum_total = sum(i * hist[i] for i in range(256))
            sum_b = 0.0
            w_b = 0
            var_max = 0.0
            threshold = 127
            for t in range(256):
                w_b += hist[t]
                if w_b == 0:
                    continue
                w_f = total - w_b
                if w_f == 0:
                    break
                sum_b += t * hist[t]
                m_b = sum_b / w_b
                m_f = (sum_total - sum_b) / w_f
                var_between = w_b * w_f * (m_b - m_f) ** 2
                if var_between > var_max:
                    var_max = var_between
                    threshold = t
            return g.point(lambda p: 255 if p > threshold else 0, mode="1")
        except Exception as e:
            logging.warning(f"OCR preprocessing failed, using raw image: {e}")
            return img

    @staticmethod
    def _render_pdf_page(file_path: str) -> ExtractionResult:
        if fitz is None:
            raise ImportError("PyMuPDF is not installed.")
        doc = fitz.open(file_path)
        page = doc[0]
        scale = DocumentExtractor.IMAGE_RENDER_SCALE
        mat = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        doc.close()
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        return ExtractionResult(content_type="image", content=b64, mime_type="image/png", method="vision")

    @staticmethod
    def _from_jpeg(file_path: str) -> ExtractionResult:
        with open(file_path, "rb") as f:
            img_bytes = f.read()
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        return ExtractionResult(content_type="image", content=b64, mime_type="image/jpeg", method="vision")


# ─────────────────────────────────────────────────────────────
# APIClient
# ─────────────────────────────────────────────────────────────

class APIClient:

    @staticmethod
    def _build_prompt(extraction: ExtractionResult) -> str:
        """Build the classification prompt. AI extracts the client name freely from
        the document — no client list in the prompt. fuzzy_match (in FileProcessor)
        maps the raw name to the authoritative list entry."""
        if extraction.content_type == "text":
            doc_section = f"Document text:\n{extraction.content}"
        else:
            doc_section = "[See attached image]"

        return (
            "You are a document classifier for a law firm.\n"
            "Your job is to identify which client this document belongs to "
            "and write a short description of what it is.\n\n"
            f"{doc_section}\n\n"
            "RULES — read carefully before responding:\n\n"
            "CLIENT IDENTIFICATION:\n"
            "- Scan the document for labels that introduce the client's name, "
            "in this priority order:\n"
            "    1. 'Client:', 'Client Name:', 'Client/Patient Name:'\n"
            "    2. 'Claimant:', 'Injured:', 'Injured Party:', 'Injured Worker:'\n"
            "    3. 'Patient:', 'Patient Name:'\n"
            "    4. 'Insured:', 'Insured Name:', 'Named Insured:'\n"
            "    5. 'Employee:' (workers comp)\n"
            "    6. 'RE:', 'Re:', 'Regarding:', 'Subject:'\n"
            "    7. Case captions (plaintiff or defendant the firm represents)\n"
            "- Do NOT identify a business, facility, clinic, hospital, insurance "
            "company, opposing party, or attorney as the client.\n"
            "- Return the client's full name exactly as it appears in the document.\n"
            "- If you cannot clearly identify the client, return NEEDS_REVIEW.\n\n"
            "DESCRIPTION:\n"
            "- 2-5 separate words with spaces between them, title case, no special characters "
            "(e.g. \"Retainer Agreement\", \"Motion to Dismiss\", \"Invoice\").\n"
            "- IMPORTANT: Always use spaces between words. Never combine words "
            "(e.g. write \"Medical Record Request\" NOT \"Medicalrecordrequest\").\n"
            "- NEVER describe a fax cover sheet or fax wrapper — describe the actual "
            "document content. If no real content is visible, use \"Incoming Document\".\n\n"
            "Return ONLY valid JSON with no extra text:\n"
            '{"client": "LAST, First", "desc": "Retainer Agreement", "confidence": "high|medium|low"}'
        )

    @staticmethod
    def classify(extraction: ExtractionResult, client_list: list, config: dict) -> dict:
        # No client list in the prompt — AI extracts the name freely from the document.
        # fuzzy_match (called in FileProcessor) maps the raw name to the authoritative list.
        prompt = APIClient._build_prompt(extraction)

        if config.get("processing", {}).get("debug_log_prompt", False):
            logging.info(f"[DEBUG PROMPT]\n{prompt}\n[END PROMPT]")

        api_cfg = config["api"]
        errors = []

        # Try OpenWebUI (OpenAI-compatible)
        try:
            raw = APIClient._call_openwebui(prompt, extraction, api_cfg)
            return APIClient._parse_response(raw)
        except Exception as e:
            errors.append(f"OpenWebUI: {e}")
            logging.warning(f"OpenWebUI API call failed: {e}")

        # Fallback: direct Ollama
        try:
            raw = APIClient._call_ollama(prompt, extraction, api_cfg)
            return APIClient._parse_response(raw)
        except Exception as e:
            errors.append(f"Ollama: {e}")
            logging.error(f"Ollama fallback also failed: {e}")

        raise ConnectionError("Both API endpoints failed.\n" + "\n".join(errors))

    @staticmethod
    def _call_openwebui(prompt: str, extraction: ExtractionResult, api_cfg: dict) -> str:
        url = api_cfg["openwebui_url"].rstrip("/") + "/api/chat/completions"
        headers = {"Content-Type": "application/json"}
        if api_cfg.get("api_key"):
            headers["Authorization"] = f"Bearer {api_cfg['api_key']}"

        if extraction.content_type == "image":
            # OpenAI vision format: content is an array. Send every page
            # the extractor produced (vision mode may return multiple).
            image_list = extraction.images or [extraction.content]
            content_parts = [
                {"type": "text", "text": prompt.replace("[See attached image]", "").strip()}
            ]
            for b64 in image_list:
                content_parts.append({
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:{extraction.mime_type};base64,{b64}"
                    },
                })
            messages = [{"role": "user", "content": content_parts}]
        else:
            messages = [{"role": "user", "content": prompt}]

        payload = {
            "model": api_cfg["model"],
            "messages": messages,
            "stream": False,
        }
        resp = requests.post(
            url, json=payload, headers=headers,
            timeout=(api_cfg["timeout_connect"], api_cfg["timeout_read"])
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]

    @staticmethod
    def _call_ollama(prompt: str, extraction: ExtractionResult, api_cfg: dict) -> str:
        url = api_cfg["ollama_url"].rstrip("/") + "/api/generate"
        payload = {
            "model": api_cfg["model"],
            "prompt": prompt,
            "format": "json",
            "stream": False,
        }
        if extraction.content_type == "image":
            payload["images"] = extraction.images or [extraction.content]

        resp = requests.post(
            url, json=payload,
            timeout=(api_cfg["timeout_connect"], api_cfg["timeout_read"])
        )
        resp.raise_for_status()
        return resp.json().get("response", "")

    @staticmethod
    def _parse_response(raw: str) -> dict:
        # Try direct parse
        try:
            data = json.loads(raw)
            if isinstance(data, dict) and "client" in data:
                return data
        except json.JSONDecodeError:
            pass
        # Extract first JSON object via regex
        match = re.search(r'\{[^{}]*\}', raw, re.DOTALL)
        if match:
            try:
                data = json.loads(match.group())
                if isinstance(data, dict):
                    return data
            except json.JSONDecodeError:
                pass
        logging.warning(f"Could not parse API response as JSON. Raw: {raw[:300]}")
        return {"client": "A-NEEDS REVIEW", "desc": "Unknown Document", "confidence": "low"}

    @staticmethod
    def test_connection(config: dict) -> tuple:
        api_cfg = config["api"]
        # Try OpenWebUI
        try:
            url = api_cfg["openwebui_url"].rstrip("/") + "/api/chat/completions"
            headers = {"Content-Type": "application/json"}
            if api_cfg.get("api_key"):
                headers["Authorization"] = f"Bearer {api_cfg['api_key']}"
            payload = {
                "model": api_cfg["model"],
                "messages": [{"role": "user", "content": "Reply with just the word: ready"}],
                "stream": False,
            }
            resp = requests.post(url, json=payload, headers=headers,
                                 timeout=(api_cfg["timeout_connect"], 30))
            resp.raise_for_status()
            return True, f"Connected via OpenWebUI  ({api_cfg['model']})"
        except Exception as e1:
            pass
        # Try Ollama direct
        try:
            url = api_cfg["ollama_url"].rstrip("/") + "/api/generate"
            payload = {
                "model": api_cfg["model"],
                "prompt": "Reply with just the word: ready",
                "stream": False,
            }
            resp = requests.post(url, json=payload,
                                 timeout=(api_cfg["timeout_connect"], 30))
            resp.raise_for_status()
            return True, f"Connected via Ollama direct  ({api_cfg['model']})"
        except Exception as e2:
            return False, f"Could not connect.\nOpenWebUI: {e1}\nOllama: {e2}"


# ─────────────────────────────────────────────────────────────
# FileProcessor
# ─────────────────────────────────────────────────────────────

class FileProcessor:

    @staticmethod
    def process_file(file_path: str, config: dict, client_list: list) -> ProcessResult:
        filename = os.path.basename(file_path)
        ext = os.path.splitext(filename)[1].lower()
        proc_cfg = config["processing"]

        # Skip already-processed files
        if proc_cfg.get("skip_already_processed") and \
                FileProcessor._already_processed(filename, client_list):
            return ProcessResult(
                original_name=filename,
                final_name=filename,
                status="skipped",
            )

        try:
            if not os.path.isfile(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
            if os.path.getsize(file_path) == 0:
                raise ValueError("File is empty (0 bytes)")

            doc_hash = FileProcessor._file_hash(file_path)

            # Decide extraction method. Vision mode is only honored when:
            #   1. The user selected it in Settings, AND
            #   2. The selected model is on the local vision allowlist.
            # Any other case silently falls back to OCR so the pipeline
            # never breaks if the user switches to a non-vision model.
            use_vision = (
                proc_cfg.get("extraction_method", "ocr") == "vision"
                and model_supports_vision(config.get("api", {}).get("model", ""))
            )

            # Extract content
            extraction = DocumentExtractor.extract(
                file_path,
                proc_cfg["max_ocr_chars"],
                proc_cfg.get("max_pages", 5),
                vision_mode=use_vision,
                max_vision_pages=proc_cfg.get("max_vision_pages", 2),
                ocr_preprocess=proc_cfg.get("ocr_preprocess", True),
            )

            # Classify via AI
            result = APIClient.classify(extraction, client_list, config)
            raw_client = result.get("client", "NEEDS_REVIEW").strip().strip("\"'")
            raw_desc = result.get("desc", "Unknown Document")
            confidence = result.get("confidence", "low")
            raw_confidence = confidence

            # First 2000 chars of the actual extracted text, for later
            # learning/dedupe features. Vision mode reads images, not text,
            # so there's nothing textual to capture there.
            extracted_text = (
                extraction.content[:2000]
                if extraction.content_type == "text" and extraction.content
                else ""
            )

            # Log what the AI returned so failures are easy to diagnose
            logging.info(f"{filename}: AI returned client='{raw_client}' confidence={confidence}")

            # Confidence gate — configurable in Settings.
            # require_high_confidence=True (default): only "high" proceeds to rename.
            # require_high_confidence=False: "medium" also proceeds (more matches, more risk).
            # "low" always goes to NEEDS_REVIEW regardless of this setting.
            require_high = proc_cfg.get("require_high_confidence", True)
            _confidence_ok = (
                confidence == "high"
                or (confidence == "medium" and not require_high)
            )
            _skip_fuzzy = (
                raw_client in ("NEEDS_REVIEW", "A-NEEDS REVIEW", "")
                or not _confidence_ok
            )
            if not _confidence_ok and raw_client not in ("NEEDS_REVIEW", "A-NEEDS REVIEW", ""):
                logging.info(
                    f"{filename}: confidence={confidence} "
                    f"({'below high threshold' if require_high else 'low'}) "
                    "— sending to NEEDS_REVIEW without matching"
                )

            threshold = proc_cfg.get("fuzzy_threshold", 0.82)
            matched = None if _skip_fuzzy else \
                ClientListManager.fuzzy_match(raw_client, client_list, threshold)

            if matched:
                logging.info(f"{filename}: fuzzy matched '{raw_client}' → '{matched}'")

            if matched:
                final_client = matched
                status = "renamed"
            else:
                final_client = "A-NEEDS REVIEW"
                status = "needs_review"

            safe_desc = FileProcessor._safe_subject(raw_desc) or "Document"

            # Safety net: if the description is still fax-related despite the prompt rule,
            # substitute a neutral fallback rather than letting it become the filename.
            if FileProcessor._desc_is_fax(safe_desc):
                logging.info(
                    f"{filename}: AI returned fax-related desc '{safe_desc}' — "
                    "substituting 'Incoming Document'"
                )
                safe_desc = "Incoming Document"

            new_name = f"{final_client} - {safe_desc}{ext}"

            # Collision avoidance
            dest_dir = os.path.dirname(file_path)
            new_name = FileProcessor._resolve_collision(dest_dir, new_name, filename)

            renamed_at = None
            if new_name != filename:
                os.rename(file_path, os.path.join(dest_dir, new_name))
                renamed_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            return ProcessResult(
                original_name=filename,
                final_name=new_name,
                status=status,
                client=final_client,
                description=safe_desc,
                confidence=confidence,
                renamed_at=renamed_at,
                extraction_method=extraction.method,
                raw_client=raw_client,
                raw_confidence=raw_confidence,
                extracted_text=extracted_text,
                doc_hash=doc_hash,
            )

        except Exception as e:
            logging.error(f"Error processing {filename}: {e}", exc_info=True)
            return ProcessResult(
                original_name=filename,
                final_name=filename,
                status="error",
                error_message=str(e),
            )

    # Phrases that indicate the AI described the fax wrapper rather than the real document
    _FAX_DESC_PATTERNS = [
        "fax", "facsimile", "telecopy", "send result", "transmission report",
        "activity report", "communication journal",
    ]

    @staticmethod
    def _desc_is_fax(desc: str) -> bool:
        lower = desc.lower()
        return any(p in lower for p in FileProcessor._FAX_DESC_PATTERNS)

    # Small joining words that stay lowercase when they aren't the first
    # token of the subject (e.g. "Reduction Request to Chiropractic Works").
    _LOWERCASE_JOINERS = {
        "to", "of", "and", "for", "in", "on", "at", "by", "the", "a", "an",
        "with", "from", "vs",
    }

    @staticmethod
    def _is_acronym(token: str) -> bool:
        """True if `token` is an ALL-CAPS 2-5 char run of letters (PPR, TTD,
        IME, MRI, EMG, NCV, PPD, TTP, DOI, DOB, PT, OT, ...). Trailing
        punctuation (e.g. a stray period) is ignored when checking."""
        core = token.strip(".,;:!?")
        return 2 <= len(core) <= 5 and core.isalpha() and core.isupper()

    @staticmethod
    def _safe_subject(text: str, preserve_acronyms: bool = True) -> str:
        # Replace slashes and underscores with spaces BEFORE stripping illegal
        # characters, so "EMG/NCV" -> "EMG NCV" and "Incoming_Document" ->
        # "Incoming Document" instead of the words being silently fused.
        text = text.replace("/", " ").replace("_", " ")
        text = ILLEGAL_CHARS_RE.sub("", text)
        text = re.sub(r"\s+", " ", text).strip()

        if not preserve_acronyms:
            # Legacy behavior: no acronym awareness at all.
            text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
            text = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", text)
            return text.title()

        # Split run-on words: insert space before uppercase letters that follow
        # a lowercase letter (e.g. "RetainerAgreement" → "Retainer Agreement").
        # Skip this when the whole token is already a recognised acronym so
        # "EMG" or "IME" are never split apart.
        def _split_run_on(token: str) -> str:
            if FileProcessor._is_acronym(token):
                return token
            token = re.sub(r"([a-z])([A-Z])", r"\1 \2", token)
            token = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", token)
            return token

        tokens = text.split(" ")
        split_tokens: List[str] = []
        for tok in tokens:
            if not tok:
                continue
            split_tokens.extend(_split_run_on(tok).split(" "))

        out_tokens = []
        for i, tok in enumerate(split_tokens):
            if FileProcessor._is_acronym(tok):
                out_tokens.append(tok)
            elif i > 0 and tok.lower() in FileProcessor._LOWERCASE_JOINERS:
                out_tokens.append(tok.lower())
            else:
                out_tokens.append(tok[:1].upper() + tok[1:].lower() if tok else tok)

        return " ".join(out_tokens).strip()

    @staticmethod
    def _resolve_collision(directory: str, filename: str, source_name: str) -> str:
        """If `filename` already exists in `directory` (and isn't the source file),
        append (1), (2), … until a free name is found."""
        if filename == source_name:
            return filename
        if not os.path.exists(os.path.join(directory, filename)):
            return filename
        base, ext = os.path.splitext(filename)
        counter = 1
        while True:
            candidate = f"{base} ({counter}){ext}"
            if not os.path.exists(os.path.join(directory, candidate)):
                return candidate
            counter += 1

    @staticmethod
    def _file_hash(path: str) -> str:
        """Return the sha256 hex digest of a file's bytes, read in chunks.
        Returns "" on any error (missing file, permission error, etc.) rather
        than raising — this is a best-effort helper for dedupe/learning
        features, not something that should ever break processing."""
        try:
            h = hashlib.sha256()
            with open(path, "rb") as f:
                for chunk in iter(lambda: f.read(65536), b""):
                    h.update(chunk)
            return h.hexdigest()
        except Exception:
            return ""

    @staticmethod
    def _already_processed(filename: str, client_list: list) -> bool:
        """True if the file already looks like 'LAST, First - Subject.ext'
        with a recognised client name at the front."""
        if filename.startswith("A-NEEDS REVIEW"):
            return False
        match = re.match(r"^(.+?) - .+\.(pdf|jpg|jpeg)$", filename, re.IGNORECASE)
        if not match:
            return False
        name_part = match.group(1)
        return ClientListManager.fuzzy_match(name_part, client_list, threshold=0.90) is not None


# ─────────────────────────────────────────────────────────────
# ProcessingEngine  (runs in a background thread)
# ─────────────────────────────────────────────────────────────

class ProcessingEngine:
    def __init__(self):
        self._stop_event = threading.Event()

    def stop(self):
        self._stop_event.set()

    def run_batch(self, config: dict, result_queue: queue.Queue):
        self._stop_event.clear()
        scandocs = config["paths"]["scandocs_folder"]
        client_list_path = config["paths"]["client_list_file"]

        try:
            client_list = ClientListManager.load(client_list_path)
            if not client_list:
                result_queue.put({
                    "type": "error",
                    "message": (
                        "The client list is empty.\n\n"
                        "Go to the Client List tab, add your clients, and save."
                    ),
                })
                result_queue.put({"type": "done"})
                return

            try:
                all_entries = os.listdir(scandocs)
            except Exception as e:
                result_queue.put({
                    "type": "error",
                    "message": f"Cannot read scandocs folder:\n{e}",
                })
                result_queue.put({"type": "done"})
                return

            files = [
                f for f in all_entries
                if os.path.isfile(os.path.join(scandocs, f))
                and os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS
            ]

            if not files:
                result_queue.put({
                    "type": "error",
                    "message": "No PDF or JPG files found in the scandocs folder.",
                })
                result_queue.put({"type": "done"})
                return

            result_queue.put({"type": "total", "count": len(files)})

            for i, filename in enumerate(files):
                if self._stop_event.is_set():
                    result_queue.put({"type": "stopped"})
                    break
                result_queue.put({
                    "type": "progress",
                    "current": i + 1,
                    "filename": filename,
                })
                result = FileProcessor.process_file(
                    os.path.join(scandocs, filename), config, client_list
                )
                result_queue.put({"type": "result", "result": result})

        except Exception as e:
            logging.error(f"Unhandled batch error: {e}", exc_info=True)
            result_queue.put({"type": "error", "message": str(e)})
        finally:
            result_queue.put({"type": "done"})


# ─────────────────────────────────────────────────────────────
# SplashScreen
# ─────────────────────────────────────────────────────────────

class SplashScreen:
    """Fun splash window shown while the main app initialises.
    Shows the logo, an animated spinner, and a stream of dad jokes."""

    DAD_JOKES = [
        ("Why don't scientists trust atoms?",
         "Because they make up everything!"),
        ("I told my wife she was drawing her eyebrows too high.",
         "She looked surprised."),
        ("I'm reading a book about anti-gravity.",
         "It's impossible to put down!"),
        ("Did you hear about the claustrophobic astronaut?",
         "He just needed a little space."),
        ("Why don't eggs tell jokes?",
         "They'd crack each other up!"),
        ("I asked the librarian if they had books about paranoia.",
         "She whispered: 'They're right behind you!'"),
        ("What do you call a fish without eyes?",
         "A fsh."),
        ("Why can't a leopard hide?",
         "Because he's always spotted!"),
        ("I used to hate facial hair…",
         "…but then it grew on me."),
        ("What do you call a factory that makes okay products?",
         "A satisfactory."),
    ]

    def __init__(self, parent: tk.Misc, on_done, primary_color: str = "#1565c0",
                 show_duration_ms: int = 7500):
        self._on_done = on_done
        self._primary = primary_color
        import random as _random
        self._joke_idx = _random.randint(0, len(self.DAD_JOKES) - 1)
        self._spinner_angle = 0
        self._dot_count = 0
        self._done_called = False

        self.win = tk.Toplevel(parent)
        self.win.overrideredirect(True)
        self.win.resizable(False, False)
        self.win.configure(bg="#ffffff")

        W, H = 520, 360
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        self.win.geometry(f"{W}x{H}+{(sw - W)//2}+{(sh - H)//2}")
        self.win.lift()
        self.win.attributes("-topmost", True)

        # DWM rounded window corners (Windows 11)
        try:
            import ctypes
            hwnd = self.win.winfo_id()
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 33, ctypes.byref(ctypes.c_int(2)), ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

        self._build(W, H)
        self._animate_spinner()
        self.win.after(400, self._show_question)
        self.win.after(show_duration_ms, self._finish)

    # ── Layout ────────────────────────────────────────────────

    def _build(self, W: int, H: int):
        primary = self._primary

        # Coloured top banner
        banner = tk.Frame(self.win, bg=primary, height=8)
        banner.pack(fill=tk.X, side=tk.TOP)

        # Coloured bottom banner
        tk.Frame(self.win, bg=primary, height=8).pack(fill=tk.X, side=tk.BOTTOM)

        # Logo
        self._logo_img = None
        png_path = os.path.join(ASSETS_DIR, "Speedy Scandocs Logo.png")
        if PILImage is not None and os.path.isfile(png_path):
            try:
                img = PILImage.open(png_path).convert("RGBA")
                bbox = img.getbbox()
                if bbox:
                    img = img.crop(bbox)
                target_h = 90
                ratio = target_h / img.height
                new_w = max(1, int(img.width * ratio))
                img = img.resize((new_w, target_h), PILImage.LANCZOS)
                bg_img = PILImage.new("RGB", (new_w, target_h), (255, 255, 255))
                bg_img.paste(img, mask=img.split()[3])
                self._logo_img = PILImageTk.PhotoImage(bg_img)
            except Exception:
                pass

        if self._logo_img:
            tk.Label(self.win, image=self._logo_img,
                     bg="#ffffff", bd=0).place(relx=0.5, y=70, anchor="center")
        else:
            tk.Label(self.win, text="Speedy Scandocs", bg="#ffffff",
                     fg=primary, font=(APP_FONT, 20, "bold")).place(
                relx=0.5, y=70, anchor="center")

        # Spinner canvas
        self._spin_canvas = tk.Canvas(
            self.win, width=46, height=46, bg="#ffffff", highlightthickness=0)
        self._spin_canvas.place(relx=0.5, y=168, anchor="center")

        # Joke label
        self._joke_var = tk.StringVar(value="")
        self._joke_lbl = tk.Label(
            self.win, textvariable=self._joke_var, bg="#ffffff",
            fg="#555555", font=(APP_FONT, 10, "italic"),
            wraplength=460, justify="center",
        )
        self._joke_lbl.place(relx=0.5, y=240, anchor="center", width=480)

        # Footer hint
        tk.Label(self.win, text="Loading, please wait…", bg="#ffffff",
                 fg="#bbbbbb", font=(APP_FONT, 8)).place(
            relx=0.5, y=320, anchor="center")

    # ── Spinner animation ─────────────────────────────────────

    def _animate_spinner(self):
        if not self.win.winfo_exists():
            return
        c = self._spin_canvas
        c.delete("all")
        a = self._spinner_angle
        p = self._primary
        c.create_arc(3, 3, 43, 43, start=a,       extent=270,
                     style="arc", outline=p,       width=5)
        c.create_arc(3, 3, 43, 43, start=a + 270, extent=90,
                     style="arc", outline="#dddddd", width=5)
        self._spinner_angle = (self._spinner_angle + 14) % 360
        self.win.after(40, self._animate_spinner)

    # ── Dad joke cycle ────────────────────────────────────────

    def _show_question(self):
        if not self.win.winfo_exists():
            return
        joke = self.DAD_JOKES[self._joke_idx % len(self.DAD_JOKES)]
        self._joke_var.set(f"{joke[0]}")
        self._dot_count = 0
        self.win.after(2200, self._animate_dots)

    def _animate_dots(self):
        if not self.win.winfo_exists():
            return
        dots = "  .  " * ((self._dot_count % 3) + 1)
        self._joke_var.set(dots)
        self._dot_count += 1
        if self._dot_count <= 5:
            self.win.after(350, self._animate_dots)
        else:
            self.win.after(200, self._show_answer)

    def _show_answer(self):
        if not self.win.winfo_exists():
            return
        joke = self.DAD_JOKES[self._joke_idx % len(self.DAD_JOKES)]
        self._joke_var.set(f"{joke[1]}")
        self._joke_idx += 1
        self.win.after(2600, self._show_question)

    # ── Close ─────────────────────────────────────────────────

    def _finish(self):
        if self._done_called:
            return
        self._done_called = True
        try:
            if self.win.winfo_exists():
                self.win.destroy()
        except Exception:
            pass
        try:
            self._on_done()
        except Exception as exc:
            logging.warning(f"SplashScreen on_done error: {exc}")


# ─────────────────────────────────────────────────────────────
# ScandocsApp  (tkinter GUI)
# ─────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────
# Auto-update (GitHub Releases)
# ─────────────────────────────────────────────────────────────

def _parse_version(tag: str) -> tuple:
    """'v1.8.0' or '1.8.0' -> (1, 8, 0). Unknown suffixes are dropped.
    Returns (0,) on a tag we can't parse so we never offer a bogus update."""
    if not tag:
        return (0,)
    s = tag.strip().lstrip("vV")
    parts = []
    for chunk in s.split("."):
        num = ""
        for ch in chunk:
            if ch.isdigit():
                num += ch
            else:
                break
        if not num:
            break
        parts.append(int(num))
    return tuple(parts) if parts else (0,)


def _pick_release_asset(release: dict) -> Optional[dict]:
    """Pick the right platform asset from a GitHub release JSON payload.
    Windows -> first .exe, Mac -> first .dmg. Returns asset dict or None."""
    assets = release.get("assets") or []
    if sys.platform == "win32":
        want = ".exe"
    elif sys.platform == "darwin":
        want = ".dmg"
    else:
        return None
    for a in assets:
        name = (a.get("name") or "").lower()
        if name.endswith(want):
            return a
    return None


def fetch_latest_release(timeout: float = 8.0) -> Optional[dict]:
    """Hit the GitHub API for the latest release. Returns parsed JSON or None.
    Silent on network failure — we don't want offline users to see errors."""
    if requests is None:
        return None
    try:
        r = requests.get(
            UPDATE_API_URL,
            headers={"Accept": "application/vnd.github+json"},
            timeout=timeout,
        )
        if r.status_code != 200:
            logging.info(f"Update check: HTTP {r.status_code}")
            return None
        return r.json()
    except Exception as e:
        logging.info(f"Update check failed: {e}")
        return None


def download_file(url: str, dest_path: str,
                  progress_cb=None, cancel_flag=None) -> bool:
    """Stream a file to dest_path. progress_cb(downloaded, total) called
    periodically. cancel_flag is a callable returning True to abort.
    Returns True on success."""
    if requests is None:
        return False
    try:
        with requests.get(url, stream=True, timeout=30) as r:
            r.raise_for_status()
            total = int(r.headers.get("Content-Length") or 0)
            downloaded = 0
            tmp = dest_path + ".part"
            with open(tmp, "wb") as f:
                for chunk in r.iter_content(chunk_size=65536):
                    if cancel_flag and cancel_flag():
                        try: os.remove(tmp)
                        except OSError: pass
                        return False
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_cb:
                            progress_cb(downloaded, total)
            os.replace(tmp, dest_path)
            return True
    except Exception as e:
        logging.warning(f"Download failed: {e}")
        return False


class ScandocsApp(ttk.Window):

    def report_callback_exception(self, exc, val, tb):
        """Tkinter normally prints callback exceptions to stderr, which is
        invisible when the app is launched without a console — silently
        swallowing bugs in button commands etc. Log them instead."""
        logging.error("Unhandled exception in a UI callback", exc_info=(exc, val, tb))

    def __init__(self):
        super().__init__(themename="litera")
        self.withdraw()               # hidden until splash finishes
        self.title(APP_TITLE)
        self.geometry("1500x1000")
        self.minsize(1100, 800)

        self._apply_app_font()

        self.config_mgr = ConfigManager()
        self.engine = ProcessingEngine()
        self._queue: queue.Queue = queue.Queue()
        self._results: list = []
        self._total_files: int = 0
        self._iid_to_result: dict = {}   # treeview item id → ProcessResult
        self._correction_buttons: dict = {}  # treeview item id → overlaid "Manual Correction" button
        self._audit_updating: bool = False  # prevent recursive checkbox callbacks
        self.fo_dest_var = tk.StringVar()        # file-mode destination folder
        self.s_file_mode_auto_var      = tk.BooleanVar(value=True)
        self.s_file_mode_manual_var    = tk.BooleanVar(value=True)
        self.s_suggest_loc_var         = tk.BooleanVar(value=False)
        self.s_suggest_parent_var      = tk.StringVar(value="")
        self.s_require_high_conf_var   = tk.BooleanVar(value=True)
        self.s_show_manual_tab_var     = tk.BooleanVar(value=False)
        self._all_clients: list = []        # full client list for combo filtering
        self._correction_iid: str = ""      # results_tree row id currently under Manual Correction
        self._pre_correction_height: Optional[int] = None  # window height before growing for the panel
        self._review_selected_file: str = ""  # last file chosen in the legacy Manual Entry list
        self._sort_col: str = ""            # treeview column currently sorted
        self._sort_reverse: bool = False    # ascending=False, descending=True
        self._file_popup = None             # currently open document preview Toplevel
        self._file_popup_canvas = None      # canvas inside the popup for refreshing

        os.makedirs(DEFAULT_REPORTS_FOLDER, exist_ok=True)

        # Show splash FIRST so the user sees something immediately while the
        # heavy UI build runs. Force a redraw so the splash paints before
        # _build_ui blocks the main thread.
        self._splash = SplashScreen(
            self, on_done=self.deiconify, primary_color=_APP_PRIMARY,
            show_duration_ms=3000,
        )
        self.update_idletasks()
        self.update()

        # Defer heavy initialisation to after_idle so the splash's event loop
        # has a chance to render and animate before we block on UI build.
        self.after_idle(self._deferred_init)

    def _deferred_init(self):
        """Run the heavy UI build after the splash has rendered."""
        self._build_ui()
        # Must run AFTER _build_ui — ttkbootstrap builds its per-color button
        # styles lazily the first time a button with that bootstyle is
        # constructed, so an earlier override would be overwritten.
        self._apply_rounded_buttons()
        self._load_settings_to_ui()
        self._refresh_client_list_tab()
        self._refresh_unnamed_count()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # Set icon after ttkbootstrap finishes its own setup to prevent it being overridden
        self.after(100, self._set_window_icon)
        self.after(200, self._check_first_run)
        # Auto-update: check on startup if enabled and 24h have passed.
        # Delay past splash so we don't race the initial UI paint.
        self.after(4000, self._maybe_check_for_updates_async)

    def _apply_app_font(self):
        """Point every named tk font, ttk style, and option-db default at
        APP_FONT so widgets built before explicit font= kwargs still pick it up."""
        import tkinter.font as tkfont
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
                     "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont",
                     "TkTooltipFont", "TkFixedFont"):
            try:
                tkfont.nametofont(name).configure(family=APP_FONT)
            except tk.TclError:
                pass
        try:
            self.style.configure(".", font=(APP_FONT, 10))
            self.style.configure("Treeview", font=(APP_FONT, 10))
            self.style.configure("Treeview.Heading", font=(APP_FONT, 10, "bold"))
            self.style.configure("TNotebook.Tab", font=(APP_FONT, 10))
            self.style.configure("TButton", font=(APP_FONT, 10))
            self.style.configure("TLabel", font=(APP_FONT, 10))
            self.style.configure("TEntry", font=(APP_FONT, 10))
            self.style.configure("TCombobox", font=(APP_FONT, 10))
            self.style.configure("TCheckbutton", font=(APP_FONT, 10))
            self.style.configure("TRadiobutton", font=(APP_FONT, 10))
            self.style.configure("TLabelframe.Label", font=(APP_FONT, 10, "bold"))
            self.style.configure("TMenubutton", font=(APP_FONT, 10))
        except Exception as e:
            logging.info(f"Could not apply ttk font: {e}")
        # Option db — covers classic tk widgets (tk.Label, tk.Button, Listbox,
        # Menu, Text, Entry) created without an explicit font= kwarg.
        self.option_add("*Font", (APP_FONT, 10))
        self.option_add("*TCombobox*Listbox.font", (APP_FONT, 10))

    def _apply_rounded_buttons(self):
        """Replace ttk button layouts with 9-slice rounded-rectangle images
        for every bootstyle color the app uses. Preserves the ttkbootstrap
        color scheme — only changes corner shape."""
        if PILImage is None:
            return
        try:
            from PIL import ImageDraw, ImageTk
        except ImportError:
            return

        radius = 10
        w, h = 220, 46
        theme = self.style.colors
        color_map = {}
        for cname in ("primary", "secondary", "success", "danger",
                      "warning", "info", "dark", "light"):
            try:
                color_map[cname] = getattr(theme, cname)
            except AttributeError:
                pass

        self._rounded_btn_imgs = []  # keep PhotoImage refs alive

        def _shade(hex_c: str, amount: float) -> str:
            hex_c = hex_c.lstrip("#")
            r = int(hex_c[0:2], 16); g = int(hex_c[2:4], 16); b = int(hex_c[4:6], 16)
            if amount >= 0:
                r = max(0, int(r * (1 - amount)))
                g = max(0, int(g * (1 - amount)))
                b = max(0, int(b * (1 - amount)))
            else:
                a = -amount
                r = min(255, int(r + (255 - r) * a))
                g = min(255, int(g + (255 - g) * a))
                b = min(255, int(b + (255 - b) * a))
            return f"#{r:02x}{g:02x}{b:02x}"

        def _make(fill=None, border=None, bw=0):
            img = PILImage.new("RGBA", (w, h), (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            kw = {}
            if fill: kw["fill"] = fill
            if border and bw:
                kw["outline"] = border
                kw["width"] = bw
            d.rounded_rectangle((0, 0, w - 1, h - 1), radius=radius, **kw)
            ph = ImageTk.PhotoImage(img)
            self._rounded_btn_imgs.append(ph)
            return ph

        def _register(style_name, el_name, normal, active, pressed, disabled,
                      fg, hover_fg=None):
            try:
                self.style.element_create(
                    el_name, "image", normal,
                    ("pressed", pressed),
                    ("active", active),
                    ("disabled", disabled),
                    border=radius, sticky="nsew",
                )
            except tk.TclError:
                return  # element already exists — safe to ignore
            self.style.layout(style_name, [
                (el_name, {"sticky": "nsew", "children": [
                    ("Button.padding", {"sticky": "nsew", "children": [
                        ("Button.label", {"sticky": "nsew"}),
                    ]}),
                ]}),
            ])
            self.style.configure(style_name, foreground=fg,
                                 font=(APP_FONT, 10),
                                 padding=(14, 7),
                                 borderwidth=0, focuscolor="",
                                 relief="flat")
            if hover_fg:
                self.style.map(style_name, foreground=[
                    ("active", hover_fg), ("pressed", hover_fg),
                ])

        # Solid + outline per color
        for name, color in color_map.items():
            fg_on_color = "#ffffff" if name != "light" else "#212529"
            _register(
                f"{name}.TButton",
                f"Rounded.{name}.button",
                _make(fill=color),
                _make(fill=_shade(color, 0.08)),
                _make(fill=_shade(color, 0.18)),
                _make(fill=_shade(color, -0.5)),
                fg=fg_on_color,
            )
            _register(
                f"{name}.Outline.TButton",
                f"Rounded.{name}.Outline.button",
                _make(fill="#ffffff", border=color, bw=2),
                _make(fill=color, border=color, bw=2),
                _make(fill=_shade(color, 0.15), border=_shade(color, 0.15), bw=2),
                _make(fill="#ffffff", border=_shade(color, -0.5), bw=2),
                fg=color,
                hover_fg="#ffffff" if name != "light" else "#212529",
            )

        # Default (un-bootstyled) button
        default_bg = color_map.get("light", "#f1f3f5")
        border_c = "#ced4da"
        _register(
            "TButton",
            "Rounded.default.button",
            _make(fill=default_bg, border=border_c, bw=1),
            _make(fill=_shade(default_bg, 0.05), border=border_c, bw=1),
            _make(fill=_shade(default_bg, 0.12), border=border_c, bw=1),
            _make(fill=default_bg, border=_shade(border_c, -0.3), bw=1),
            fg="#212529",
        )

    # ── UI Construction ───────────────────────────────────────

    def _set_window_icon(self):
        """Set the title bar and taskbar icon to the GDJ logo."""
        png_path = os.path.join(ASSETS_DIR, "GDJ Logo.png")
        ico_path = os.path.join(ASSETS_DIR, "GDJ Logo.ico")

        # Generate the .ico from PNG if it doesn't exist yet (Windows)
        if not os.path.isfile(ico_path) and PILImage is not None and os.path.isfile(png_path):
            try:
                src = PILImage.open(png_path).convert("RGBA")
                sizes = [256, 128, 64, 48, 32, 16]
                frames = [src.resize((s, s), PILImage.LANCZOS) for s in sizes]
                frames[0].save(ico_path, format="ICO",
                               append_images=frames[1:],
                               sizes=[(s, s) for s in sizes])
            except Exception as e:
                logging.warning(f"Could not create icon file: {e}")

        if sys.platform == "darwin":
            # macOS: use iconphoto with a PNG — iconbitmap(.ico) does not work
            if PILImage is not None and os.path.isfile(png_path):
                try:
                    img = PILImage.open(png_path).convert("RGBA")
                    self._app_icon_photo = PILImageTk.PhotoImage(img)
                    self.iconphoto(True, self._app_icon_photo)
                except Exception as e:
                    logging.warning(f"iconphoto (macOS) failed: {e}")
        else:
            # Windows/Linux: use iconbitmap with .ico
            if os.path.isfile(ico_path):
                try:
                    self.iconbitmap(ico_path)
                except Exception as e:
                    logging.warning(f"iconbitmap failed: {e}")

    def _build_ui(self):
        self._build_header()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self._build_process_tab()
        self._build_review_tab()   # legacy Manual Entry tab — attached/detached via Settings toggle
        self._build_clients_tab()
        self._build_settings_tab()
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _on_tab_changed(self, _evt=None):
        # Refresh each tab when it becomes active so it always reflects the
        # current scandocs folder / client list on disk.
        try:
            current = self.notebook.nametowidget(self.notebook.select())
        except Exception:
            return
        if getattr(self, "_process_tab", None) is current:
            self._refresh_unnamed_count()
        elif getattr(self, "_review_tab", None) is current:
            self._refresh_review_tab()
        elif getattr(self, "_clients_tab", None) is current:
            self._refresh_client_list_tab()

    def _refresh_unnamed_count(self):
        """Count files in the scandocs folder that don't yet match the
        'LAST, First - Subject.ext' format, and show it on the Auto-Process tab."""
        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        if not scandocs or not os.path.isdir(scandocs):
            self.unnamed_count_var.set("Scandocs folder not configured or not found.")
            return
        client_list_path = self.config_mgr.config["paths"]["client_list_file"]
        client_list = ClientListManager.load(client_list_path)
        try:
            unnamed = [
                f for f in os.listdir(scandocs)
                if os.path.isfile(os.path.join(scandocs, f))
                and os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS
                and not FileProcessor._already_processed(f, client_list)
            ]
        except OSError as e:
            self.unnamed_count_var.set(f"Could not read scandocs folder: {e}")
            return
        n = len(unnamed)
        if n == 0:
            self.unnamed_count_var.set("No unnamed files in scandocs folder.")
        elif n == 1:
            self.unnamed_count_var.set("1 unnamed file ready to process.")
        else:
            self.unnamed_count_var.set(f"{n} unnamed files ready to process.")

    def _build_header(self):
        """Header bar: Speedy Scandocs logo PNG, right-aligned."""
        header = tk.Frame(self, bg="#ffffff", height=80)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        # Hairline separator beneath the header (palette-coloured)
        self._header_sep = tk.Frame(self, bg="#dee2e6", height=2)
        self._header_sep.pack(fill=tk.X)

        self._header_logo = None
        png_path = os.path.join(ASSETS_DIR, "Speedy Scandocs Logo.png")
        if PILImage is not None and os.path.isfile(png_path):
            try:
                img = PILImage.open(png_path).convert("RGBA")
                # Trim surrounding whitespace so the logo fills the space cleanly
                bbox = img.getbbox()
                if bbox:
                    img = img.crop(bbox)
                # Scale to fit inside the header with a small top/bottom margin
                target_h = 64
                ratio = target_h / img.height
                new_w = max(1, int(img.width * ratio))
                img = img.resize((new_w, target_h), PILImage.LANCZOS)
                # Composite onto white so RGBA looks clean on the white header
                bg = PILImage.new("RGB", (new_w, target_h), (255, 255, 255))
                bg.paste(img, mask=img.split()[3])
                self._header_logo = PILImageTk.PhotoImage(bg)
            except Exception as e:
                logging.warning(f"Could not load header logo: {e}")

        if self._header_logo:
            lbl = tk.Label(header, image=self._header_logo, bg="#ffffff", bd=0)
            lbl.place(relx=1.0, rely=0.5, anchor="e", x=-16)
        else:
            # Fallback text if image unavailable
            tk.Label(header, text="Speedy Scandocs", bg="#ffffff",
                     fg="#212529", font=(APP_FONT, 15, "bold")).place(
                relx=1.0, rely=0.5, anchor="e", x=-16)

    # ── Tab 1: Process ────────────────────────────────────────

    def _build_process_tab(self):
        self.style.configure("AutoTab.TFrame", background="#e3f2fd")
        tab = ttk.Frame(self.notebook, style="AutoTab.TFrame")
        self._process_tab = tab
        self._process_tab_frame = tab
        self.notebook.add(tab, text="  Process  ")
        # Accent bar
        tk.Frame(tab, bg=_APP_PRIMARY, height=6).pack(fill=tk.X)

        # Button row
        btn_row = ttk.Frame(tab)
        btn_row.pack(fill=tk.X, padx=10, pady=(10, 4))
        self.btn_process = ttk.Button(
            btn_row, text="Auto-Process Documents", command=self._start_processing,
            bootstyle="primary",
        )
        self.btn_process.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_stop = ttk.Button(
            btn_row, text="Stop", command=self._stop_processing,
            state=tk.DISABLED, bootstyle="danger-outline",
        )
        self.btn_stop.pack(side=tk.LEFT)
        right_frame = ttk.Frame(btn_row)
        right_frame.pack(side=tk.RIGHT)
        self.btn_open_report = ttk.Button(
            right_frame, text="Open Report", command=self._open_report,
            bootstyle="primary-outline",
        )
        self.btn_open_report.pack(anchor="e")
        self.btn_submit_audit = ttk.Button(
            right_frame, text="Submit Audit", command=self._submit_audit,
            bootstyle="primary",
        )
        # btn_submit_audit visibility is toggled by _apply_audit_mode
        ttk.Label(
            right_frame, text="Reports Saved Automatically",
            font=(APP_FONT, 7), foreground="gray",
        ).pack(anchor="e")

        # Unnamed-files indicator (refreshes when this tab is selected)
        self.unnamed_count_var = tk.StringVar(value="")
        ttk.Label(
            tab, textvariable=self.unnamed_count_var,
            font=(APP_FONT, 10, "bold"), foreground="#1565c0",
            anchor="w",
        ).pack(fill=tk.X, padx=10, pady=(2, 0))

        # Progress
        prog_frame = ttk.Frame(tab)
        prog_frame.pack(fill=tk.X, padx=10, pady=(0, 4))
        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(
            prog_frame, variable=self.progress_var, maximum=100
        ).pack(fill=tk.X)
        self.status_var = tk.StringVar(
            value="Ready — configure Settings then press Auto-Process Documents."
        )
        ttk.Label(prog_frame, textvariable=self.status_var, anchor="w").pack(
            fill=tk.X, pady=(2, 0)
        )

        # Results table
        cols = ("audited", "original", "new_name", "status", "client",
                "new_location", "correction")
        col_cfg = {
            "audited":      ("✓",              32),
            "original":     ("Original File",  180),
            "new_name":     ("New Name",        240),
            "status":       ("Status",           75),
            "client":       ("Client",          170),
            "new_location": ("New Location",    130),
            "correction":   ("",                 160),
        }
        tree_frame = ttk.Frame(tab)
        self._tree_frame = tree_frame   # stable pack anchor — always packed, unlike audit_panel
        # tree_frame is packed AFTER the bottom panels so it fills remaining space

        self.results_tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings", selectmode="extended"
        )
        # Columns that support click-to-sort
        _sortable = {"original", "new_name", "status", "client"}
        for col in cols:
            label, width = col_cfg[col]
            if col in _sortable:
                self.results_tree.heading(
                    col, text=label,
                    command=lambda c=col: self._sort_treeview(c))
            else:
                self.results_tree.heading(col, text=label)
            self.results_tree.column(
                col, width=width,
                minwidth=32 if col == "audited" else 50,
                anchor="center" if col in ("audited", "correction") else "w",
            )
        self.results_tree.tag_configure("renamed",      background="#d4edda")
        self.results_tree.tag_configure("needs_review", background="#fff3cd")
        self.results_tree.tag_configure("error",        background="#f8d7da")
        self.results_tree.tag_configure("skipped",      background="#e2e3e5")
        self.results_tree.tag_configure("audited",      background="#fddcb0")
        self.results_tree.tag_configure("moved",        background="#cce5ff")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=self.results_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.results_tree.xview)

        def _yscroll_set(*args):
            vsb.set(*args)
            self._reposition_all_correction_buttons()

        def _xscroll_set(*args):
            hsb.set(*args)
            self._reposition_all_correction_buttons()

        self.results_tree.configure(yscrollcommand=_yscroll_set, xscrollcommand=_xscroll_set)
        vsb.pack(side=tk.RIGHT,  fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.results_tree.pack(fill=tk.BOTH, expand=True)
        self.results_tree.bind("<<TreeviewSelect>>", self._on_result_select)
        self.results_tree.bind("<space>",           self._on_tree_return)
        self.results_tree.bind("<Double-Button-1>", self._on_tree_double_click)
        self.results_tree.bind("<Left>",   self._audit_prev)
        self.results_tree.bind("<Right>",  self._audit_next)
        self.results_tree.bind("<Configure>", lambda e: self._reposition_all_correction_buttons())
        # Mousewheel scroll on the treeview
        _tree_scroll = lambda e: self.results_tree.yview_scroll(
            -1 * (e.delta // 120) if e.delta else (-1 if e.num == 4 else 1), "units")
        self._bind_mousewheel(self.results_tree, _tree_scroll)

        # Pack tree now — bottom panels pack after (side=BOTTOM) so they always show
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 0))

        # ── Audit panel ──────────────────────────────────────
        self.audit_panel = ttk.LabelFrame(tab, text="Audit Mode")
        self.audit_panel.pack(fill=tk.X, padx=10, pady=(6, 0))
        audit_outer = self.audit_panel

        # Row 1: filename + Open File button
        file_row = ttk.Frame(audit_outer)
        file_row.pack(fill=tk.X, padx=8, pady=(6, 2))
        self.audit_file_label = ttk.Label(
            file_row, text="Select a row above to audit it.",
            foreground="gray", anchor="w",
        )
        self.audit_file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.audit_open_btn = ttk.Button(
            file_row, text="Open File", bootstyle="dark-outline",
            command=self._audit_open_file, state=tk.DISABLED,
        )
        self.audit_open_btn.pack(side=tk.RIGHT, padx=(8, 0))
        self.audit_next_btn = ttk.Button(
            file_row, text="Next →", bootstyle="dark-outline",
            command=self._audit_next, state=tk.DISABLED,
        )
        self.audit_next_btn.pack(side=tk.RIGHT, padx=(4, 0))
        self.audit_prev_btn = ttk.Button(
            file_row, text="← Prev", bootstyle="dark-outline",
            command=self._audit_prev, state=tk.DISABLED,
        )
        self.audit_prev_btn.pack(side=tk.RIGHT, padx=(4, 0))

        # Checkboxes — two rows so all are always visible
        self.audit_correct_var       = tk.BooleanVar()
        self.audit_wrong_client_var  = tk.BooleanVar()
        self.audit_bad_desc_var      = tk.BooleanVar()
        self.audit_failed_client_var = tk.BooleanVar()
        self.audit_should_flag_var   = tk.BooleanVar()

        def _make_chk(parent, text, var, flag, fg="#212529"):
            return tk.Checkbutton(
                parent, text=text, variable=var,
                command=lambda: self._on_audit_check(flag),
                fg=fg, bg="#f8f9fa",
                activeforeground=fg, activebackground="#f8f9fa",
                selectcolor="#ffffff",
                disabledforeground="#adb5bd",
                font=(APP_FONT, 11),
                padx=8, pady=6,
                bd=2,
                state=tk.DISABLED,
            )

        # Row 2a: positive check
        chk_row1 = ttk.Frame(audit_outer)
        chk_row1.pack(fill=tk.X, padx=8, pady=(2, 0))
        self.audit_correct_chk = _make_chk(
            chk_row1, "✓  Correct", self.audit_correct_var, "correct", fg="#1a6e31")
        self.audit_correct_chk.pack(side=tk.LEFT)

        # Row 2b: problem flags
        chk_row2 = ttk.Frame(audit_outer)
        chk_row2.pack(fill=tk.X, padx=8, pady=(2, 8))

        self.audit_wrong_client_chk = _make_chk(
            chk_row2, "Wrong client name", self.audit_wrong_client_var, "wrong_client")
        self.audit_wrong_client_chk.pack(side=tk.LEFT, padx=(0, 24))

        self.audit_bad_desc_chk = _make_chk(
            chk_row2, "Bad description", self.audit_bad_desc_var, "bad_description")
        self.audit_bad_desc_chk.pack(side=tk.LEFT, padx=(0, 24))

        self.audit_failed_client_chk = _make_chk(
            chk_row2, "Failed to identify client", self.audit_failed_client_var, "failed_client")
        self.audit_failed_client_chk.pack(side=tk.LEFT, padx=(0, 24))

        self.audit_should_flag_chk = _make_chk(
            chk_row2, "Should have been flagged for review", self.audit_should_flag_var, "should_review")
        self.audit_should_flag_chk.pack(side=tk.LEFT)

        # Row 2c: rename hint — shown when any problem flag is active
        self.audit_rename_hint_var = tk.StringVar(value="")
        self.audit_rename_hint_lbl = tk.Label(
            audit_outer,
            textvariable=self.audit_rename_hint_var,
            font=(APP_FONT, 8, "italic"),
            fg="#777777",
            bg="#f8f9fa",
            anchor="w",
            wraplength=0,   # no wrapping — single line
        )
        self.audit_rename_hint_lbl.pack(fill=tk.X, padx=8, pady=(0, 6))

        # ── File Operations panel ─────────────────────────────
        self.file_ops_panel = ttk.LabelFrame(tab, text="File Mode — Move Files")
        self.file_ops_panel.pack(fill=tk.X, padx=10, pady=(6, 10))

        # Row 1: destination folder + browse + apply-to-selected
        dest_row = ttk.Frame(self.file_ops_panel)
        dest_row.pack(fill=tk.X, padx=8, pady=(8, 4))
        ttk.Label(dest_row, text="Destination Folder:").pack(side=tk.LEFT)
        ttk.Entry(dest_row, textvariable=self.fo_dest_var, width=40).pack(
            side=tk.LEFT, padx=(8, 4), fill=tk.X, expand=True)
        ttk.Button(
            dest_row, text="Browse",
            command=self._fo_browse_dest,
            bootstyle="dark-outline",
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            dest_row, text="Apply to Selected",
            command=self._fo_apply_to_selected,
            bootstyle="dark-outline",
        ).pack(side=tk.LEFT)

        # Row 2: commit button + status
        action_row = ttk.Frame(self.file_ops_panel)
        action_row.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Button(
            action_row, text="Move Files",
            command=self._fo_move_all,
            bootstyle="primary",
        ).pack(side=tk.LEFT)
        self.fo_status_var = tk.StringVar(value="")
        ttk.Label(action_row, textvariable=self.fo_status_var,
                  foreground="gray").pack(side=tk.LEFT, padx=14)

        # ── Manual Correction panel (hidden until a row's "Manual Correction"
        #    cell is clicked) ─────────────────────────────────
        self.correction_panel = ttk.LabelFrame(tab, text="Manual Correction")
        # Not packed here — _open_manual_correction() shows it, _hide_manual_correction() hides it.

        corr = self.correction_panel

        corr_file_row = ttk.Frame(corr)
        corr_file_row.pack(fill=tk.X, padx=8, pady=(8, 2))
        self.corr_file_var = tk.StringVar(value="")
        ttk.Label(corr_file_row, textvariable=self.corr_file_var,
                  font=(APP_FONT, 10, "bold")).pack(side=tk.LEFT)

        # Large, unmistakable preview button — spacebar types a space in this
        # panel's fields instead of toggling the viewer, so this button is the
        # clear way to see the document while correcting it.
        ttk.Button(
            corr, text="🔍  Open File Preview",
            command=self._corr_open_preview,
            bootstyle="info", padding=(16, 12),
        ).pack(fill=tk.X, padx=8, pady=(2, 10))

        corr_fields = ttk.Frame(corr)
        corr_fields.pack(fill=tk.X, padx=8, pady=(0, 4))
        corr_fields.columnconfigure(1, weight=1)

        # Row 0: Client typing entry
        ttk.Label(corr_fields, text="Client:").grid(
            row=0, column=0, padx=(0, 8), pady=(0, 2), sticky="nw"
        )
        self.corr_client_var = tk.StringVar()
        corr_client_entry = ttk.Entry(
            corr_fields, textvariable=self.corr_client_var, width=40
        )
        corr_client_entry.grid(row=0, column=1, pady=(0, 2), sticky="w")
        corr_client_entry.bind("<KeyRelease>", self._filter_client_combo)
        corr_client_entry.bind("<Return>", lambda e: self._corr_next())

        # Row 1: Always-visible scrollable client list — shrinks as user types
        client_lb_frame = ttk.Frame(corr_fields)
        client_lb_frame.grid(row=1, column=1, pady=(0, 4), sticky="ew")
        self.corr_client_listbox = tk.Listbox(
            client_lb_frame, height=6,
            font=(APP_FONT, 9),
            selectmode=tk.SINGLE,
            exportselection=False,
            activestyle="none",
        )
        cl_lb_vsb = ttk.Scrollbar(client_lb_frame, orient="vertical",
                                   command=self.corr_client_listbox.yview)
        self.corr_client_listbox.configure(yscrollcommand=cl_lb_vsb.set)
        cl_lb_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.corr_client_listbox.pack(fill=tk.X, expand=True)
        self.corr_client_listbox.bind(
            "<<ListboxSelect>>", self._on_client_listbox_select)
        self.corr_client_listbox.bind(
            "<Return>", lambda e: self._corr_next())

        # Row 2: Subject
        ttk.Label(corr_fields, text="Subject:").grid(
            row=2, column=0, padx=(0, 8), pady=6, sticky="w"
        )
        self.corr_subject_var = tk.StringVar()
        corr_subject_entry = ttk.Entry(
            corr_fields, textvariable=self.corr_subject_var, width=40)
        corr_subject_entry.grid(row=2, column=1, pady=6, sticky="w")
        corr_subject_entry.bind("<Return>", lambda e: self._corr_next())

        # Row 3: Destination folder (shown only when File Mode is on for Manual Correction)
        self.corr_dest_label = ttk.Label(corr_fields, text="Move to folder:")
        self.corr_dest_label.grid(row=3, column=0, padx=(0, 8), pady=6, sticky="w")
        self.corr_dest_label.grid_remove()
        corr_dest_inner = ttk.Frame(corr_fields)
        corr_dest_inner.grid(row=3, column=1, pady=6, sticky="ew")
        corr_dest_inner.grid_remove()
        ttk.Entry(corr_dest_inner, textvariable=self.fo_dest_var).pack(
            side=tk.LEFT, padx=(0, 4), fill=tk.X, expand=True)
        ttk.Button(corr_dest_inner, text="Browse",
                   command=self._fo_browse_dest,
                   bootstyle="dark-outline").pack(side=tk.LEFT)
        self._corr_dest_widgets = [self.corr_dest_label, corr_dest_inner]

        # Row 4: Close / Next buttons
        corr_btn_row = ttk.Frame(corr)
        corr_btn_row.pack(fill=tk.X, padx=8, pady=(4, 10))
        ttk.Button(corr_btn_row, text="Next →",
                   command=self._corr_next,
                   bootstyle="primary").pack(side=tk.RIGHT)
        ttk.Button(corr_btn_row, text="Close",
                   command=self._corr_close,
                   bootstyle="dark-outline").pack(side=tk.RIGHT, padx=(0, 6))

    # ── Legacy Manual Entry tab (Settings: "Show Manual Entry tab") ──
    # Built once at startup but only attached to the notebook (via
    # _apply_manual_tab_visibility) when the Settings toggle is on.

    def _build_review_tab(self):
        tab = ttk.Frame(self.notebook)
        self._review_tab = tab
        self._review_tab_frame = tab
        tk.Frame(tab, bg=_APP_PRIMARY, height=5).pack(fill=tk.X)

        top = ttk.Frame(tab)
        top.pack(fill=tk.X, padx=10, pady=(10, 0))
        self.review_count_var = tk.StringVar(value="Files awaiting review: 0")
        ttk.Label(top, textvariable=self.review_count_var,
                  font=(APP_FONT, 10, "bold")).pack(side=tk.LEFT)
        ttk.Button(top, text="Refresh", command=self._refresh_review_tab,
                   bootstyle="dark-outline").pack(side=tk.RIGHT)
        ttk.Button(top, text="Next →", command=self._review_next,
                   bootstyle="dark-outline").pack(side=tk.RIGHT, padx=(4, 6))
        ttk.Button(top, text="← Prev", command=self._review_prev,
                   bootstyle="dark-outline").pack(side=tk.RIGHT, padx=(0, 4))

        # File list
        list_lf = ttk.LabelFrame(tab, text="Files Awaiting Review")
        list_lf.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        self.review_listbox = tk.Listbox(list_lf, selectmode=tk.SINGLE,
                                         font=(APP_FONT, 10), exportselection=False)
        rv_sb = ttk.Scrollbar(list_lf, orient="vertical", command=self.review_listbox.yview)
        self.review_listbox.configure(yscrollcommand=rv_sb.set)
        rv_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.review_listbox.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self.review_listbox.bind("<<ListboxSelect>>", self._on_review_select)
        self.review_listbox.bind("<space>",           self._on_review_return)
        self.review_listbox.bind("<Double-Button-1>", self._on_review_double_click)
        self.review_listbox.bind("<Left>",   self._review_prev)
        self.review_listbox.bind("<Right>",  self._review_next)
        # Up/Down already navigate the Listbox natively and fire <<ListboxSelect>>.
        _rv_scroll = lambda e: self.review_listbox.yview_scroll(
            -1 * (e.delta // 120) if e.delta else (-1 if e.num == 4 else 1), "units")
        self._bind_mousewheel(self.review_listbox, _rv_scroll)

        # Assignment panel
        assign_lf = ttk.LabelFrame(tab, text="Assign Selected File")
        assign_lf.pack(fill=tk.X, padx=10, pady=(0, 10))
        assign_lf.columnconfigure(1, weight=1)

        ttk.Label(assign_lf, text="Open file:").grid(
            row=0, column=0, padx=8, pady=6, sticky="w"
        )
        ttk.Button(assign_lf, text="Open in Viewer",
                   command=self._open_review_file,
                   bootstyle="dark-outline").grid(
            row=0, column=1, padx=4, pady=6, sticky="w"
        )

        # Row 1: Client typing entry
        ttk.Label(assign_lf, text="Client:").grid(
            row=1, column=0, padx=8, pady=(6, 2), sticky="nw"
        )
        self.review_client_var = tk.StringVar()
        review_client_entry = ttk.Entry(
            assign_lf, textvariable=self.review_client_var, width=40
        )
        review_client_entry.grid(row=1, column=1, padx=4, pady=(6, 2), sticky="w")
        review_client_entry.bind("<KeyRelease>", self._filter_review_client_combo)
        review_client_entry.bind("<Return>", lambda e: self._assign_review_file())

        # Row 2: Always-visible scrollable client list — shrinks as user types
        client_lb_frame = ttk.Frame(assign_lf)
        client_lb_frame.grid(row=2, column=1, padx=4, pady=(0, 4), sticky="ew")
        self.review_client_listbox = tk.Listbox(
            client_lb_frame, height=7,
            font=(APP_FONT, 9),
            selectmode=tk.SINGLE,
            exportselection=False,
            activestyle="none",
        )
        cl_lb_vsb = ttk.Scrollbar(client_lb_frame, orient="vertical",
                                   command=self.review_client_listbox.yview)
        self.review_client_listbox.configure(yscrollcommand=cl_lb_vsb.set)
        cl_lb_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.review_client_listbox.pack(fill=tk.X, expand=True)
        self.review_client_listbox.bind(
            "<<ListboxSelect>>", self._on_review_client_listbox_select)
        self.review_client_listbox.bind(
            "<Return>", lambda e: self._assign_review_file())

        # Row 3: Subject
        ttk.Label(assign_lf, text="Subject:").grid(
            row=3, column=0, padx=8, pady=6, sticky="w"
        )
        self.review_subject_var = tk.StringVar()
        review_subject_entry = ttk.Entry(
            assign_lf, textvariable=self.review_subject_var, width=40)
        review_subject_entry.grid(row=3, column=1, padx=4, pady=6, sticky="w")
        review_subject_entry.bind("<Return>", lambda e: self._assign_review_file())

        # Row 4: Destination folder (shown only when File Mode is on)
        self.review_dest_label = ttk.Label(assign_lf, text="Move to folder:")
        self.review_dest_label.grid(row=4, column=0, padx=8, pady=6, sticky="w")
        self.review_dest_label.grid_remove()
        review_dest_inner = ttk.Frame(assign_lf)
        review_dest_inner.grid(row=4, column=1, padx=4, pady=6, sticky="ew")
        review_dest_inner.grid_remove()
        ttk.Entry(review_dest_inner, textvariable=self.fo_dest_var).pack(
            side=tk.LEFT, padx=(0, 4), fill=tk.X, expand=True)
        ttk.Button(review_dest_inner, text="Browse",
                   command=self._fo_browse_dest,
                   bootstyle="dark-outline").pack(side=tk.LEFT)
        self._review_dest_widgets = [self.review_dest_label, review_dest_inner]

        # Row 5: Submit button
        _btn_assign = ttk.Button(assign_lf, text="Assign, Rename, & Move",
                   command=self._assign_review_file,
                   bootstyle="primary")
        _btn_assign.grid(
            row=5, column=1, padx=4, pady=(2, 8), sticky="w"
        )

    # ── Tab 2: Client List ────────────────────────────────────

    def _build_clients_tab(self):
        tab = ttk.Frame(self.notebook)
        self._clients_tab = tab
        self.notebook.add(tab, text="  Client List  ")

        ttk.Label(
            tab,
            text="Format: LAST, First   (e.g.  GARCIA, Maria  |  REYES, Carlos A)",
            foreground="gray",
        ).pack(anchor="w", padx=10, pady=(10, 2))

        # List + scrollbar
        list_frame = ttk.Frame(tab)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)
        self.client_listbox = tk.Listbox(
            list_frame, selectmode=tk.SINGLE, font=(APP_FONT, 10)
        )
        cl_sb = ttk.Scrollbar(list_frame, orient="vertical",
                               command=self.client_listbox.yview)
        self.client_listbox.configure(yscrollcommand=cl_sb.set)
        cl_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.client_listbox.pack(fill=tk.BOTH, expand=True)
        _cl_scroll = lambda e: self.client_listbox.yview_scroll(
            -1 * (e.delta // 120) if e.delta else (-1 if e.num == 4 else 1), "units")
        self._bind_mousewheel(self.client_listbox, _cl_scroll)

        # Add / Remove
        edit_frame = ttk.Frame(tab)
        edit_frame.pack(fill=tk.X, padx=10, pady=(0, 2))
        self.new_client_var = tk.StringVar()
        entry = ttk.Entry(edit_frame, textvariable=self.new_client_var, width=32)
        entry.pack(side=tk.LEFT, padx=(0, 6))
        entry.bind("<Return>", lambda _: self._add_client())
        entry.bind("<KeyRelease>", lambda _: self._on_client_entry_key())
        _btn_add_cl = ttk.Button(edit_frame, text="Add Client",
                   command=self._add_client,
                   bootstyle="primary-outline")
        _btn_add_cl.pack(side=tk.LEFT, padx=(0, 6))
        _btn_rm_cl = ttk.Button(edit_frame, text="Remove Selected",
                   command=self._remove_client,
                   bootstyle="danger-outline")
        _btn_rm_cl.pack(side=tk.LEFT)

        # Dynamic "Add X to list" suggestion — shown when typed name is not yet in the list
        self._add_suggestion_lbl = tk.Label(
            tab, text="", fg="#1565c0", bg="#ffffff",
            font=(APP_FONT, 9, "underline"), cursor="hand2", anchor="w",
        )
        self._add_suggestion_lbl.pack(fill=tk.X, padx=12, pady=(0, 2))
        self._add_suggestion_lbl.bind("<Button-1>", lambda _: self._add_client())

        # Save row
        save_frame = ttk.Frame(tab)
        save_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        _btn_save_cl = ttk.Button(save_frame, text="Save Client List",
                   command=self._save_client_list,
                   bootstyle="primary")
        _btn_save_cl.pack(side=tk.LEFT)
        self.client_status_var = tk.StringVar(value="")
        ttk.Label(save_frame, textvariable=self.client_status_var,
                  foreground="green").pack(side=tk.LEFT, padx=10)

    # ── Tab 3: Settings ───────────────────────────────────────

    def _build_settings_tab(self):
        tab = ttk.Frame(self.notebook)
        self._settings_tab_frame = tab
        self.notebook.add(tab, text="  Settings  ")

        # ── Pinned footer (always visible at the bottom) ──────
        # Pack BEFORE the scroll canvas so pack reserves its space first.
        _footer_frame = ttk.Frame(tab, padding=(20, 8, 20, 10))
        _footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Separator(tab, orient="horizontal").pack(side=tk.BOTTOM, fill=tk.X)

        # ── Scrollable wrapper ────────────────────────────────
        _scroll_canvas = tk.Canvas(tab, highlightthickness=0)
        _vsb = ttk.Scrollbar(tab, orient="vertical", command=_scroll_canvas.yview)
        _scroll_canvas.configure(yscrollcommand=_vsb.set)
        _vsb.pack(side=tk.RIGHT, fill=tk.Y)
        _scroll_canvas.pack(fill=tk.BOTH, expand=True)

        outer = ttk.Frame(_scroll_canvas)
        _win_id = _scroll_canvas.create_window((0, 0), window=outer, anchor="nw")

        def _on_inner_configure(e):
            _scroll_canvas.configure(scrollregion=_scroll_canvas.bbox("all"))
        outer.bind("<Configure>", _on_inner_configure)

        def _on_canvas_configure(e):
            _scroll_canvas.itemconfig(_win_id, width=e.width)
        _scroll_canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(e):
            _scroll_canvas.yview_scroll(
                -1 * (e.delta // 120) if e.delta else (-1 if e.num == 4 else 1), "units")
        # Bind scroll recursively after the tab is fully built (so all children exist)
        self.after(200, lambda: ScandocsApp._bind_mousewheel(tab, _on_mousewheel))

        # ── Inner padding frame (same variable name 'outer' the rest of
        #    _build_settings_tab uses, so nothing below changes) ──────
        outer = ttk.Frame(outer, padding=(20, 15, 20, 15))
        outer.pack(fill=tk.BOTH, expand=True)

        def add_row(parent, row_idx, label_text, str_var,
                    browse_dir=False, browse_file=False, masked=False, info_msg=None):
            ttk.Label(parent, text=label_text).grid(
                row=row_idx, column=0, sticky="w", pady=5
            )
            kw = {"textvariable": str_var, "width": 46}
            if masked:
                kw["show"] = "*"
            entry = ttk.Entry(parent, **kw)
            entry.grid(row=row_idx, column=1, sticky="ew", padx=(8, 4))
            if browse_dir:
                _b = ttk.Button(
                    parent, text="Browse",
                    command=lambda v=str_var: self._browse_dir(v),
                    bootstyle="primary-outline",
                )
                _b.grid(row=row_idx, column=2, padx=(0, 4))
            elif browse_file:
                _b = ttk.Button(
                    parent, text="Browse",
                    command=lambda v=str_var: self._browse_file(v),
                    bootstyle="primary-outline",
                )
                _b.grid(row=row_idx, column=2, padx=(0, 4))
            if info_msg:
                _q = ttk.Button(
                    parent, text="?", width=2,
                    command=lambda msg=info_msg: messagebox.showinfo("Help", msg),
                    bootstyle="primary-outline",
                )
                _q.grid(row=row_idx, column=3, padx=(0, 4))

        # Paths
        paths_lf = ttk.LabelFrame(outer, text="Paths")
        paths_lf.pack(fill=tk.X, pady=(0, 10))
        paths_lf.columnconfigure(1, weight=1)
        self.s_scandocs_var = tk.StringVar()
        self.s_client_list_var = tk.StringVar()
        add_row(paths_lf, 0, "Scandocs Folder:", self.s_scandocs_var, browse_dir=True)
        add_row(paths_lf, 1, "Client List File:", self.s_client_list_var, browse_file=True)

        # API
        api_lf = ttk.LabelFrame(outer, text="API Settings")
        api_lf.pack(fill=tk.X, pady=(0, 10))
        api_lf.columnconfigure(1, weight=1)
        self.s_owui_url_var = tk.StringVar()
        self.s_ollama_url_var = tk.StringVar()
        self.s_model_var = tk.StringVar()
        self.s_api_key_var = tk.StringVar()
        add_row(api_lf, 0, "OpenWebUI URL:", self.s_owui_url_var)
        add_row(api_lf, 1, "Ollama URL:", self.s_ollama_url_var)

        # Model row — editable Combobox + Refresh button
        _FALLBACK_MODELS = [
            "llama3.2-vision", "llama3.2", "llama3.1", "llama3",
            "mistral", "mixtral", "phi3", "gemma2", "qwen2",
            "deepseek-r1", "llava", "bakllava",
        ]
        ttk.Label(api_lf, text="Model:").grid(row=2, column=0, sticky="w", pady=5)
        self.s_model_combo = ttk.Combobox(
            api_lf, textvariable=self.s_model_var,
            values=_FALLBACK_MODELS, width=44,
            state="readonly",
        )
        self.s_model_combo.grid(row=2, column=1, sticky="ew", padx=(8, 4))
        _disable_combobox_scroll(self.s_model_combo)
        ttk.Button(
            api_lf, text="⟳", width=3,
            command=self._refresh_models,
            bootstyle="primary-outline",
        ).grid(row=2, column=2, padx=(0, 4))

        add_row(api_lf, 3, "API Key (optional):", self.s_api_key_var, masked=True,
                info_msg=(
                    "API Key\n\n"
                    "Only required if your OpenWebUI instance has authentication enabled.\n\n"
                    "To find your key: in OpenWebUI go to Settings → Account → "
                    "API Keys, then generate or copy an existing key.\n\n"
                    "Leave blank if your server does not require authentication "
                    "(typical for local setups)."
                ))

        # Processing
        proc_lf = ttk.LabelFrame(outer, text="Processing")
        proc_lf.pack(fill=tk.X, pady=(0, 10))
        proc_lf.columnconfigure(1, weight=1)
        self.s_threshold_var = tk.StringVar()
        self.s_max_chars_var = tk.StringVar()
        self.s_max_pages_var = tk.StringVar()
        add_row(proc_lf, 0, "Fuzzy Match Threshold (0.0 – 1.0):", self.s_threshold_var,
                info_msg=(
                    "Fuzzy Match Threshold\n\n"
                    "Controls how closely a client name found in a document must match "
                    "a name in your client list before it is accepted.\n\n"
                    "1.0 = exact match only\n"
                    "0.85 = allows small differences (recommended)\n"
                    "0.70 = more lenient — may cause false matches\n\n"
                    "If the tool is failing to recognise clients whose names appear "
                    "slightly differently in documents (e.g. missing accents, "
                    "abbreviations), try lowering this value slightly."
                ))
        add_row(proc_lf, 1, "Max OCR Characters:", self.s_max_chars_var,
                info_msg=(
                    "Max OCR Characters\n\n"
                    "The maximum number of characters of extracted text that will be "
                    "sent to the AI for classification.\n\n"
                    "Higher values give the AI more context but increase the time each "
                    "file takes to process.\n\n"
                    "2000–4000 is usually sufficient for identifying the client and "
                    "document type from the first page or two of a document."
                ))
        add_row(proc_lf, 2, "Max Pages Per Document:", self.s_max_pages_var,
                info_msg=(
                    "Max Pages Per Document\n\n"
                    "The maximum number of pages that will be read from each PDF.\n\n"
                    "Keeping this low (5–10) prevents very large documents from "
                    "slowing down the batch. Client names and document types are "
                    "almost always found within the first few pages."
                ))
        ttk.Checkbutton(
            proc_lf,
            text="Require high confidence — only rename when AI is confident (recommended)",
            variable=self.s_require_high_conf_var,
        ).grid(row=3, column=0, columnspan=3, sticky="w", padx=8, pady=(4, 0))
        ttk.Label(
            proc_lf,
            text="  When unchecked, medium-confidence results are also renamed (more matches, higher false-positive risk)",
            font=(APP_FONT, 8), foreground="gray",
        ).grid(row=4, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 6))

        self.s_ocr_preprocess_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            proc_lf,
            text="Preprocess scans before OCR — improves accuracy on messy/low-DPI scans (recommended)",
            variable=self.s_ocr_preprocess_var,
        ).grid(row=8, column=0, columnspan=3, sticky="w", padx=8, pady=(4, 0))
        ttk.Button(
            proc_lf, text="?", width=2,
            command=lambda: messagebox.showinfo("Preprocess Scans Before OCR",
                "Preprocess Scans Before OCR\n\n"
                "When enabled, each page image is cleaned up before it reaches "
                "Tesseract:\n\n"
                "  • Upscale — small/low-DPI scans are enlarged toward 300 DPI\n"
                "  • Autocontrast — normalizes faded or shadowed pages\n"
                "  • Binarize — converts the image to pure black and white so "
                "ink stands out from paper texture\n\n"
                "This usually improves OCR accuracy on real-world scans but adds "
                "a small amount of time per page. Turn it off if you notice "
                "worse results on already-clean digital PDFs."),
            bootstyle="primary-outline",
        ).grid(row=8, column=3, padx=(0, 4), pady=(4, 0))
        ttk.Label(
            proc_lf,
            text="  Upscales, normalizes contrast, and binarizes each page image — adds ~100–300ms per page",
            font=(APP_FONT, 8), foreground="gray",
        ).grid(row=9, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 6))

        # ── Text Extraction Method (OCR vs Vision) ────────────────
        self.s_extraction_method_var = tk.StringVar()
        self.s_max_vision_pages_var = tk.StringVar()

        ttk.Label(proc_lf, text="Text Extraction Method:").grid(
            row=5, column=0, sticky="w", pady=5
        )
        self.s_method_combo = ttk.Combobox(
            proc_lf,
            textvariable=self.s_extraction_method_var,
            values=["Use OCR", "Use Vision Model"],
            state="readonly",
            width=44,
        )
        self.s_method_combo.grid(row=5, column=1, sticky="ew", padx=(8, 4))
        _disable_combobox_scroll(self.s_method_combo)
        self.s_method_combo.bind(
            "<<ComboboxSelected>>", lambda _e: self._apply_extraction_method_ui()
        )
        ttk.Button(
            proc_lf, text="?", width=2,
            command=lambda: messagebox.showinfo("Text Extraction Method",
                "Text Extraction Method\n\n"
                "Use OCR (default): extracts text from documents with Tesseract OCR "
                "and sends the text to the AI model. Fast and reliable for clean scans.\n\n"
                "Use Vision Model: sends the first few page images directly to a "
                "vision-capable model (Llama 3.2 Vision, Gemma 4). Can help on "
                "low-quality scans, stamps, and letterheads, but is slower.\n\n"
                "Vision Model is only selectable when the currently chosen model "
                "supports vision. Cloud variants are disabled to keep client "
                "documents on-device."),
            bootstyle="primary-outline",
        ).grid(row=5, column=3, padx=(0, 4))

        ttk.Label(proc_lf, text="Max Vision Pages:").grid(
            row=6, column=0, sticky="w", pady=5
        )
        self.s_vision_pages_entry = ttk.Entry(
            proc_lf, textvariable=self.s_max_vision_pages_var, width=46
        )
        self.s_vision_pages_entry.grid(row=6, column=1, sticky="ew", padx=(8, 4))
        ttk.Button(
            proc_lf, text="?", width=2,
            command=lambda: messagebox.showinfo("Max Vision Pages",
                "Max Vision Pages\n\n"
                "How many pages from each PDF are sent to the vision model. "
                "Each page is an image, so this is much more expensive than OCR "
                "— keep it low (1–2 is usually plenty to find the client name).\n\n"
                "Only applies when 'Use Vision Model' is selected."),
            bootstyle="primary-outline",
        ).grid(row=6, column=3, padx=(0, 4))

        self._method_hint_lbl = ttk.Label(
            proc_lf, text="", font=(APP_FONT, 8), foreground="gray",
        )
        self._method_hint_lbl.grid(row=7, column=0, columnspan=4,
                                    sticky="w", padx=8, pady=(0, 6))

        # Re-evaluate vision availability when the user changes the model
        self.s_model_combo.bind(
            "<<ComboboxSelected>>", lambda _e: self._apply_extraction_method_ui()
        )
        self.s_model_var.trace_add(
            "write", lambda *_: self._apply_extraction_method_ui()
        )

        # Reports
        rep_lf = ttk.LabelFrame(outer, text="Reports")
        rep_lf.pack(fill=tk.X, pady=(0, 10))
        rep_lf.columnconfigure(1, weight=1)
        self.s_report_folder_var = tk.StringVar()
        add_row(rep_lf, 0, "Report Folder:", self.s_report_folder_var, browse_dir=True)
        self.s_auto_save_var = tk.BooleanVar()
        ttk.Checkbutton(
            rep_lf, text="Auto-save report when batch completes",
            variable=self.s_auto_save_var,
        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 6))

        self.s_audit_mode_var = tk.BooleanVar()
        ttk.Checkbutton(
            rep_lf,
            text="Audit Mode — prompt employee to review each result before closing",
            variable=self.s_audit_mode_var,
            command=self._apply_audit_mode,
        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 4))
        _q_audit = ttk.Button(
            rep_lf, text="?", width=2,
            command=lambda: messagebox.showinfo("Audit Mode",
                "Audit Mode\n\n"
                "When enabled, an Audit Mode panel appears below the results table.\n\n"
                "For each document, mark one of:\n"
                "  ✓ Correct — the rename was accurate\n"
                "  Wrong Client Name — Submit Audit will rename it to A-NEEDS REVIEW\n"
                "  Bad Description — Submit Audit will change the description to 'Scanned Document'\n"
                "  Failed to Identify Client — tool could not find the client\n"
                "  Should Have Been Flagged — should have been sent to Manual Correction\n\n"
                "Click 'Submit Audit' to apply all corrections and save the report."),
            bootstyle="primary-outline",
        )
        _q_audit.grid(row=2, column=3, padx=(0, 4), pady=(0, 4), sticky="w")

        self.s_file_mode_var = tk.BooleanVar()
        ttk.Checkbutton(
            rep_lf,
            text="File Mode — allow moving renamed files to a destination folder",
            variable=self.s_file_mode_var,
            command=self._apply_file_mode,
        ).grid(row=3, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 2))
        _q_file = ttk.Button(
            rep_lf, text="?", width=2,
            command=lambda: messagebox.showinfo("File Mode",
                "File Mode\n\n"
                "When enabled, a 'File Mode — Move Files' panel appears on the "
                "selected tabs.\n\n"
                "Use 'Apply to Selected' to stage destination folders for one or more "
                "files, then 'Move Files' to commit all staged moves.\n\n"
                "You can enable File Mode independently for the Process tab and "
                "the Manual Correction panel."),
            bootstyle="primary-outline",
        )
        _q_file.grid(row=3, column=3, padx=(0, 4), pady=(0, 2), sticky="w")

        # Sub-checkboxes: per-tab enable
        sub_frame = ttk.Frame(rep_lf)
        sub_frame.grid(row=4, column=0, columnspan=4, sticky="w", padx=28, pady=(0, 8))
        self._fm_auto_chk = ttk.Checkbutton(
            sub_frame, text="Enable in Process tab",
            variable=self.s_file_mode_auto_var,
            command=self._apply_file_mode,
        )
        self._fm_auto_chk.pack(side=tk.LEFT, padx=(0, 16))
        self._fm_manual_chk = ttk.Checkbutton(
            sub_frame, text="Enable in Manual Correction panel",
            variable=self.s_file_mode_manual_var,
            command=self._apply_file_mode,
        )
        self._fm_manual_chk.pack(side=tk.LEFT)
        self._file_mode_sub_frame = sub_frame

        ttk.Checkbutton(
            rep_lf,
            text="Show Manual Entry tab (legacy) — a separate tab for assigning "
                 "files without going through Auto-Process",
            variable=self.s_show_manual_tab_var,
            command=self._apply_manual_tab_visibility,
        ).grid(row=6, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 6))

        # ── Suggest Location sub-option ───────────────────────────
        self._fm_suggest_frame = ttk.Frame(rep_lf)
        self._fm_suggest_frame.grid(row=5, column=0, columnspan=4, sticky="w",
                                    padx=28, pady=(0, 4))

        self._fm_suggest_chk = ttk.Checkbutton(
            self._fm_suggest_frame,
            text="Suggest Location — automatically find the client's folder",
            state=tk.DISABLED,
        )
        self._fm_suggest_chk.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(
            self._fm_suggest_frame,
            text="Coming Soon",
            foreground="#999999",
            font=(APP_FONT, 8, "italic"),
        ).pack(side=tk.LEFT)

        # ── Auto-commit stub (Coming Soon) ────────────────────────
        self._fm_autocommit_frame = ttk.Frame(rep_lf)
        self._fm_autocommit_frame.grid(row=7, column=0, columnspan=4, sticky="w",
                                       padx=28, pady=(0, 8))
        _ac_chk = ttk.Checkbutton(
            self._fm_autocommit_frame,
            text="Auto-commit file moves",
            state=tk.DISABLED,
        )
        _ac_chk.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(
            self._fm_autocommit_frame,
            text="Coming Soon — The tool will file documents automatically without human oversight.",
            foreground="#999999",
            font=(APP_FONT, 8, "italic"),
        ).pack(side=tk.LEFT)

        # Buttons + status — pinned to footer so they're always visible
        btn_row = ttk.Frame(_footer_frame)
        btn_row.pack(fill=tk.X)
        self._btn_test_conn = ttk.Button(btn_row, text="Test API Connection",
                   command=self._test_connection,
                   bootstyle="primary-outline")
        self._btn_test_conn.pack(side=tk.LEFT, padx=(0, 10))
        self._btn_save_settings = ttk.Button(btn_row, text="Save Settings",
                   command=self._save_settings,
                   bootstyle="primary")
        self._btn_save_settings.pack(side=tk.LEFT)
        self.conn_status_var = tk.StringVar(value="")
        self.conn_label = ttk.Label(btn_row, textvariable=self.conn_status_var)
        self.conn_label.pack(side=tk.LEFT, padx=14)

        # Version + update controls — right-aligned on the same footer row
        self._btn_check_updates = ttk.Button(
            btn_row, text="Check for Updates",
            command=lambda: self._check_for_updates_async(silent=False),
            bootstyle="secondary-outline",
        )
        self._btn_check_updates.pack(side=tk.RIGHT)
        ttk.Label(
            btn_row, text=f"v{APP_VERSION}",
            foreground="#888888",
            font=(APP_FONT, 9),
        ).pack(side=tk.RIGHT, padx=(0, 12))

    # ── Settings helpers ──────────────────────────────────────

    def _load_settings_to_ui(self):
        cfg = self.config_mgr.config
        self.s_scandocs_var.set(cfg["paths"]["scandocs_folder"])
        self.s_client_list_var.set(cfg["paths"]["client_list_file"])
        self.s_owui_url_var.set(cfg["api"]["openwebui_url"])
        self.s_ollama_url_var.set(cfg["api"]["ollama_url"])
        self.s_model_var.set(cfg["api"]["model"])
        self.s_api_key_var.set(cfg["api"]["api_key"])
        self.s_threshold_var.set(str(cfg["processing"]["fuzzy_threshold"]))
        self.s_max_chars_var.set(str(cfg["processing"]["max_ocr_chars"]))
        self.s_report_folder_var.set(
            cfg["reports"].get("report_folder", DEFAULT_REPORTS_FOLDER)
        )
        self.s_auto_save_var.set(cfg["reports"].get("auto_save", True))
        self.s_audit_mode_var.set(cfg["processing"].get("audit_mode", True))
        self.s_file_mode_var.set(cfg["processing"].get("file_mode", False))
        self.s_file_mode_auto_var.set(cfg["processing"].get("file_mode_auto", True))
        self.s_file_mode_manual_var.set(cfg["processing"].get("file_mode_manual", True))
        self.s_show_manual_tab_var.set(cfg["processing"].get("show_manual_entry_tab", False))
        self.fo_dest_var.set(cfg["processing"].get("file_mode_destination", ""))
        self.s_suggest_loc_var.set(cfg["processing"].get("suggest_location_enabled", False))
        self.s_suggest_parent_var.set(cfg["processing"].get("suggest_location_parent_folder", ""))
        self.s_max_pages_var.set(str(cfg["processing"].get("max_pages", 5)))
        self.s_require_high_conf_var.set(cfg["processing"].get("require_high_confidence", True))
        self.s_ocr_preprocess_var.set(cfg["processing"].get("ocr_preprocess", True))
        # Extraction method
        _method = cfg["processing"].get("extraction_method", "ocr")
        self.s_extraction_method_var.set(
            "Use Vision Model" if _method == "vision" else "Use OCR"
        )
        self.s_max_vision_pages_var.set(
            str(cfg["processing"].get("max_vision_pages", 2))
        )
        self.after(0, self._apply_extraction_method_ui)
        self.after(0, self._apply_audit_mode)
        self.after(0, self._apply_file_mode)
        self.after(0, self._apply_manual_tab_visibility)
        self.after(400, self._apply_round_styling)
        # Populate model list in background (won't block startup)
        self.after(300, self._refresh_models)

    def _save_settings(self):
        try:
            threshold = float(self.s_threshold_var.get())
            if not (0.0 <= threshold <= 1.0):
                raise ValueError("Fuzzy match threshold must be between 0.0 and 1.0.")
            max_chars = int(self.s_max_chars_var.get())
            if max_chars < 100:
                raise ValueError("Max OCR characters must be at least 100.")
            max_pages = int(self.s_max_pages_var.get())
            if max_pages < 1:
                raise ValueError("Max pages must be at least 1.")
            try:
                max_vision_pages = int(self.s_max_vision_pages_var.get())
            except ValueError:
                raise ValueError("Max Vision Pages must be a whole number.")
            if max_vision_pages < 1:
                raise ValueError("Max Vision Pages must be at least 1.")
        except ValueError as e:
            messagebox.showerror("Invalid Value", str(e))
            return

        cfg = self.config_mgr.config
        cfg["paths"]["scandocs_folder"]   = self.s_scandocs_var.get().strip()
        cfg["paths"]["client_list_file"]  = self.s_client_list_var.get().strip()
        cfg["api"]["openwebui_url"]        = self.s_owui_url_var.get().strip()
        cfg["api"]["ollama_url"]           = self.s_ollama_url_var.get().strip()
        cfg["api"]["model"]                = self.s_model_var.get().strip()
        cfg["api"]["api_key"]              = self.s_api_key_var.get().strip()
        cfg["processing"]["fuzzy_threshold"] = threshold
        cfg["processing"]["max_ocr_chars"]   = max_chars
        cfg["processing"]["max_pages"]        = max_pages
        cfg["reports"]["report_folder"] = (
            self.s_report_folder_var.get().strip() or DEFAULT_REPORTS_FOLDER
        )
        cfg["reports"]["auto_save"] = self.s_auto_save_var.get()
        cfg["processing"]["audit_mode"]             = self.s_audit_mode_var.get()
        cfg["processing"]["file_mode"]              = self.s_file_mode_var.get()
        cfg["processing"]["file_mode_auto"]         = self.s_file_mode_auto_var.get()
        cfg["processing"]["file_mode_manual"]       = self.s_file_mode_manual_var.get()
        cfg["processing"]["show_manual_entry_tab"]  = self.s_show_manual_tab_var.get()
        cfg["processing"]["file_mode_destination"]        = self.fo_dest_var.get().strip()
        cfg["processing"]["suggest_location_enabled"]     = self.s_suggest_loc_var.get()
        cfg["processing"]["suggest_location_parent_folder"] = self.s_suggest_parent_var.get().strip()
        cfg["processing"]["require_high_confidence"]       = self.s_require_high_conf_var.get()
        cfg["processing"]["ocr_preprocess"]                = self.s_ocr_preprocess_var.get()
        # Extraction method — force OCR if the selected model can't do vision
        _method_label = self.s_extraction_method_var.get()
        _method = "vision" if _method_label == "Use Vision Model" else "ocr"
        if _method == "vision" and not model_supports_vision(cfg["api"]["model"]):
            _method = "ocr"
        cfg["processing"]["extraction_method"] = _method
        cfg["processing"]["max_vision_pages"]  = max_vision_pages
        self.config_mgr.save(cfg)
        self._apply_audit_mode()
        self._apply_file_mode()
        self._refresh_client_list_tab()
        self._refresh_unnamed_count()
        messagebox.showinfo("Saved", "Settings saved successfully.")

    def _browse_dir(self, var: tk.StringVar):
        init = var.get() or SCRIPT_DIR
        d = filedialog.askdirectory(title="Select Folder", initialdir=init)
        if d:
            var.set(os.path.normpath(d))

    def _browse_file(self, var: tk.StringVar):
        init = os.path.dirname(var.get()) if var.get() else SCRIPT_DIR
        f = filedialog.askopenfilename(
            title="Select File",
            initialdir=init,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if f:
            var.set(os.path.normpath(f))

    def _refresh_models(self):
        """Fetch available models from Ollama and populate the Model combobox."""
        _FALLBACK_MODELS = [
            "llama3.2-vision", "llama3.2", "llama3.1", "llama3",
            "mistral", "mixtral", "phi3", "gemma2", "qwen2",
            "deepseek-r1", "llava", "bakllava",
        ]
        ollama_url = self.s_ollama_url_var.get().strip() or "http://localhost:11434"

        def _run():
            try:
                import requests as _req
                resp = _req.get(
                    f"{ollama_url.rstrip('/')}/api/tags",
                    timeout=5,
                )
                resp.raise_for_status()
                data = resp.json()
                models = sorted(m["name"] for m in data.get("models", []))
                if not models:
                    raise ValueError("Empty model list")
            except Exception:
                models = _FALLBACK_MODELS
            self.after(0, lambda: self._apply_model_list(models))

        threading.Thread(target=_run, daemon=True).start()

    def _apply_model_list(self, models: list):
        current = self.s_model_var.get().strip()
        self.s_model_combo["values"] = models
        # Keep the current selection if it's still valid, otherwise set first entry
        if current not in models and models:
            # Only overwrite if the field is blank or was a fallback default
            pass  # leave whatever the user typed / saved
        if not current and models:
            self.s_model_var.set(models[0])

    def _test_connection(self):
        self.conn_status_var.set("Testing…")
        self.conn_label.config(foreground="gray")
        self.update_idletasks()
        # Build a temporary config from the current (unsaved) field values
        cfg = ConfigManager._deep_copy(self.config_mgr.config)
        cfg["api"].update({
            "openwebui_url": self.s_owui_url_var.get().strip(),
            "ollama_url":    self.s_ollama_url_var.get().strip(),
            "model":         self.s_model_var.get().strip(),
            "api_key":       self.s_api_key_var.get().strip(),
        })

        def _run():
            ok, msg = APIClient.test_connection(cfg)
            self.after(0, lambda: self._show_conn_result(ok, msg))

        threading.Thread(target=_run, daemon=True).start()

    def _show_conn_result(self, ok: bool, msg: str):
        self.conn_status_var.set(msg)
        self.conn_label.config(foreground="green" if ok else "red")

    # ── Client list tab helpers ───────────────────────────────

    def _on_client_entry_key(self):
        """Show/hide the Add suggestion label as the user types."""
        typed = self.new_client_var.get().strip()
        if not typed:
            self._add_suggestion_lbl.config(text="")
            return
        existing = [c.lower() for c in self.client_listbox.get(0, tk.END)]
        if typed.lower() in existing:
            self._add_suggestion_lbl.config(text="")
        else:
            self._add_suggestion_lbl.config(text=f'  + Add "{typed}" to client list')

    def _refresh_client_list_tab(self):
        path = self.config_mgr.config["paths"]["client_list_file"]
        clients = ClientListManager.load(path)
        self.client_listbox.delete(0, tk.END)
        for c in clients:
            self.client_listbox.insert(tk.END, c)

    def _add_client(self):
        name = self.new_client_var.get().strip()
        if not name:
            return
        if not ClientListManager.is_valid_format(name):
            # Completely invalid (no comma, etc.) — hard warning, don't add
            messagebox.showwarning(
                "Invalid Format",
                f'"{name}" is not in LAST, First format.\n\nExample: GARCIA, Maria',
            )
            return
        # Check for the expected ALL-CAPS LAST, Title First convention
        _convention_ok = bool(re.match(
            r"^[A-Z][A-Z\-\'\.\s]+,\s+[A-Z][a-z]", name.strip()
        ))
        if not _convention_ok:
            if not messagebox.askyesno(
                "Format Warning",
                "The client name format doesn't match the expected format (LAST, First).\n\n"
                "Expected: ALL-CAPS last name, comma, Title-case first name\n"
                f'Example: GARCIA, Maria\n\nYou entered: "{name}"\n\n'
                "Would you like to proceed anyway?"
            ):
                return
        existing = list(self.client_listbox.get(0, tk.END))
        if name in existing:
            messagebox.showinfo("Duplicate", f'"{name}" is already in the list.')
            return
        existing.append(name)
        existing.sort()
        self.client_listbox.delete(0, tk.END)
        for item in existing:
            self.client_listbox.insert(tk.END, item)
        self.new_client_var.set("")
        self._add_suggestion_lbl.config(text="")
        self.client_status_var.set("Unsaved changes")

    def _remove_client(self):
        sel = self.client_listbox.curselection()
        if not sel:
            return
        name = self.client_listbox.get(sel[0])
        if messagebox.askyesno("Remove", f'Remove "{name}" from the list?'):
            self.client_listbox.delete(sel[0])
            self.client_status_var.set("Unsaved changes")

    def _save_client_list(self):
        path = self.config_mgr.config["paths"]["client_list_file"]
        if not path:
            messagebox.showerror(
                "Not Configured",
                "Client list file path is not set.\nConfigure it in Settings first.",
            )
            return
        clients = list(self.client_listbox.get(0, tk.END))
        try:
            ClientListManager.save(path, clients)
            self.client_status_var.set(f"Saved {len(clients)} client(s)")
            self.after(3000, lambda: self.client_status_var.set(""))
        except Exception as e:
            messagebox.showerror("Save Failed", str(e))

    # ── Manual Correction helpers ──────────────────────────────

    def _refresh_review_tab(self):
        """Reload the client-list cache used by the Manual Correction panel's
        client entry/listbox, and — if the legacy Manual Entry tab is enabled —
        rescan the scandocs folder for files still awaiting review. Called
        after processing finishes, after a correction is applied, and after
        an audit is submitted."""
        client_list_path = self.config_mgr.config["paths"]["client_list_file"]
        self._all_clients = ClientListManager.load(client_list_path)
        if hasattr(self, "corr_client_listbox"):
            self._filter_client_combo()

        if not hasattr(self, "review_listbox"):
            return
        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        self._review_selected_file = ""
        self.review_listbox.delete(0, tk.END)
        if not os.path.isdir(scandocs):
            self.review_count_var.set("Scandocs folder not found")
        else:
            review_files = sorted(
                f for f in os.listdir(scandocs)
                if os.path.isfile(os.path.join(scandocs, f))
                and os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS
                and not FileProcessor._already_processed(f, self._all_clients)
            )
            for f in review_files:
                self.review_listbox.insert(tk.END, f)
            self.review_count_var.set(f"Files awaiting review: {len(review_files)}")
        self._filter_review_client_combo()

    def _apply_manual_tab_visibility(self):
        """Attach or detach the legacy Manual Entry tab based on the Settings toggle."""
        show = self.s_show_manual_tab_var.get()
        present = str(self._review_tab_frame) in self.notebook.tabs()
        if show and not present:
            self.notebook.insert(1, self._review_tab_frame, text="  Manual Entry  ")
            self._refresh_review_tab()
        elif not show and present:
            self.notebook.forget(self._review_tab_frame)

    def _on_review_select(self, _event):
        sel = self.review_listbox.curselection()
        if not sel:
            return
        filename = self.review_listbox.get(sel[0])
        self._review_selected_file = filename   # remember even if focus moves away
        # Pre-fill the subject field: extract whatever comes after " - " if present,
        # otherwise use the bare filename (without extension) as a starting point.
        m = re.match(r"^.+? - (.+)\.(pdf|jpg|jpeg)$", filename, re.IGNORECASE)
        if m:
            self.review_subject_var.set(m.group(1))
        else:
            self.review_subject_var.set(os.path.splitext(filename)[0])

        # If the preview popup is open, refresh it with the newly selected file
        if self._file_popup is not None and self._file_popup.winfo_exists():
            path = os.path.join(
                self.config_mgr.config["paths"]["scandocs_folder"], filename)
            if os.path.isfile(path):
                self._refresh_popup_content(path)

    def _open_review_file(self):
        sel = self.review_listbox.curselection()
        filename = (self.review_listbox.get(sel[0]) if sel
                    else self._review_selected_file)
        if not filename:
            messagebox.showinfo("No Selection", "Please select a file from the list first.")
            return
        path = os.path.join(
            self.config_mgr.config["paths"]["scandocs_folder"], filename
        )
        if os.path.isfile(path):
            self._open_file_popup(path)
        else:
            messagebox.showerror("File Not Found", f"{filename} no longer exists.")
            self._refresh_review_tab()

    def _assign_review_file(self):
        # Use the stored selection — it persists even when focus is on the Assign fields
        sel = self.review_listbox.curselection()
        filename = (self.review_listbox.get(sel[0]) if sel
                    else self._review_selected_file)
        if not filename:
            messagebox.showinfo("No Selection", "Please select a file from the list first.")
            return
        client  = self.review_client_var.get().strip()
        subject = self.review_subject_var.get().strip()

        if not client:
            messagebox.showwarning("Missing Client", "Please select a client from the dropdown.")
            return
        if not subject:
            messagebox.showwarning("Missing Subject", "Please enter a subject for the document.")
            return

        ext      = os.path.splitext(filename)[1].lower()
        safe_sub = FileProcessor._safe_subject(subject) or "Document"
        new_name = f"{client} - {safe_sub}{ext}"

        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        new_name = FileProcessor._resolve_collision(scandocs, new_name, filename)
        src = os.path.join(scandocs, filename)
        dst = os.path.join(scandocs, new_name)
        try:
            os.rename(src, dst)
            # If file mode is on and a destination is set, move the renamed file there
            if self.s_file_mode_var.get():
                dest = self.fo_dest_var.get().strip()
                if dest and os.path.isdir(dest):
                    import shutil
                    final_name = new_name
                    final_dst = os.path.join(dest, final_name)
                    counter = 1
                    base2, ext2 = os.path.splitext(final_name)
                    while os.path.exists(final_dst):
                        final_name = f"{base2} ({counter}){ext2}"
                        final_dst = os.path.join(dest, final_name)
                        counter += 1
                    shutil.move(dst, final_dst)
                    messagebox.showinfo("Success",
                        f"Renamed and moved to:\n{os.path.basename(dest)}")
                else:
                    messagebox.showinfo("Success", f"Renamed to:\n{new_name}")
            else:
                messagebox.showinfo("Success", f"Renamed to:\n{new_name}")
            self._refresh_review_tab()
            self.review_client_var.set("")
            self.review_subject_var.set("")
        except Exception as e:
            messagebox.showerror("Rename Failed", str(e))

    def _review_prev(self, _event=None):
        """Move to the previous file in the legacy Manual Entry list."""
        n = self.review_listbox.size()
        if n == 0:
            return
        sel = self.review_listbox.curselection()
        idx = (sel[0] - 1) if sel else n - 1
        idx = max(idx, 0)
        self.review_listbox.selection_clear(0, tk.END)
        self.review_listbox.selection_set(idx)
        self.review_listbox.see(idx)
        self._on_review_select(None)

    def _review_next(self, _event=None):
        """Move to the next file in the legacy Manual Entry list."""
        n = self.review_listbox.size()
        if n == 0:
            return
        sel = self.review_listbox.curselection()
        idx = (sel[0] + 1) if sel else 0
        idx = min(idx, n - 1)
        self.review_listbox.selection_clear(0, tk.END)
        self.review_listbox.selection_set(idx)
        self.review_listbox.see(idx)
        self._on_review_select(None)

    def _on_review_return(self, _event=None):
        """Spacebar on the legacy Manual Entry list: toggle the file viewer."""
        sel = self.review_listbox.curselection()
        filename = (self.review_listbox.get(sel[0]) if sel
                    else self._review_selected_file)
        if not filename:
            return "break"
        path = os.path.join(
            self.config_mgr.config["paths"]["scandocs_folder"], filename)
        if os.path.isfile(path):
            self._toggle_file_popup(path)
        return "break"

    def _on_review_double_click(self, _event=None):
        """Double-click on the legacy Manual Entry list: always open the file viewer."""
        sel = self.review_listbox.curselection()
        filename = (self.review_listbox.get(sel[0]) if sel
                    else self._review_selected_file)
        if not filename:
            return "break"
        path = os.path.join(
            self.config_mgr.config["paths"]["scandocs_folder"], filename)
        if os.path.isfile(path):
            self._open_file_popup(path)
        return "break"

    def _filter_review_client_combo(self, _event=None):
        """Filter the legacy Manual Entry tab's client list as the user types."""
        typed = self.review_client_var.get().lower()
        self.review_client_listbox.delete(0, tk.END)
        for name in self._all_clients:
            if not typed or typed in name.lower():
                self.review_client_listbox.insert(tk.END, name)

    def _on_review_client_listbox_select(self, _event=None):
        """Clicking a name in the legacy client list copies it to the entry field."""
        sel = self.review_client_listbox.curselection()
        if sel:
            self.review_client_var.set(self.review_client_listbox.get(sel[0]))

    def _open_manual_correction(self, iid: str):
        """Show the Manual Correction panel, pre-filled with this row's
        auto-process result so the employee only has to fix what's wrong."""
        result = self._iid_to_result.get(iid)
        if not result:
            return

        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        filename = result.final_name
        if not os.path.isfile(os.path.join(scandocs, filename)):
            filename = result.original_name

        self._correction_iid = iid
        self.results_tree.selection_set(iid)
        self.results_tree.see(iid)
        self.corr_file_var.set(f"Correcting:  {filename}")

        # Populate with the auto-process info first — if the client/subject
        # split is already right, the employee only needs to touch the field
        # that's wrong (often just the file name itself).
        m = re.match(r"^(.+?) - (.+)\.(pdf|jpg|jpeg)$", filename, re.IGNORECASE)
        if m:
            self.corr_client_var.set(m.group(1))
            self.corr_subject_var.set(m.group(2))
        else:
            self.corr_client_var.set(result.client or "")
            self.corr_subject_var.set(result.description or os.path.splitext(filename)[0])

        self._filter_client_combo()

        if not self.correction_panel.winfo_ismapped():
            self.correction_panel.pack(fill=tk.X, padx=10, pady=(6, 4),
                                        after=self._tree_frame)
            self._grow_window_for_correction_panel()

    def _hide_manual_correction(self):
        self.correction_panel.pack_forget()
        self._correction_iid = ""
        self._shrink_window_after_correction_panel()

    def _grow_window_for_correction_panel(self):
        """Grow the main window so the newly-shown Manual Correction panel
        doesn't just squeeze the results table — all its fields stay visible."""
        if self.state() == "zoomed":
            return   # already fullscreen; the tree just shrinks to make room
        self.update_idletasks()
        panel_h = self.correction_panel.winfo_reqheight()
        if self._pre_correction_height is None:
            self._pre_correction_height = self.winfo_height()
        screen_h = self.winfo_screenheight()
        new_h = min(self._pre_correction_height + panel_h, screen_h - 60)
        self.geometry(f"{self.winfo_width()}x{new_h}+{self.winfo_x()}+{self.winfo_y()}")

    def _shrink_window_after_correction_panel(self):
        """Restore the window to its pre-panel height, if we grew it."""
        if self._pre_correction_height is None or self.state() == "zoomed":
            self._pre_correction_height = None
            return
        self.geometry(
            f"{self.winfo_width()}x{self._pre_correction_height}"
            f"+{self.winfo_x()}+{self.winfo_y()}")
        self._pre_correction_height = None

    def _corr_open_preview(self):
        """The Manual Correction panel's big 'Open File Preview' button."""
        result = self._iid_to_result.get(self._correction_iid)
        if not result:
            return
        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        path = os.path.join(scandocs, result.final_name)
        if not os.path.isfile(path):
            path = os.path.join(scandocs, result.original_name)
        if os.path.isfile(path):
            self._open_file_popup(path)
        else:
            messagebox.showwarning("File Not Found",
                f"Could not locate the file:\n{result.final_name}")

    def _corr_commit(self) -> bool:
        """Apply the Manual Correction panel's fields to the file under
        correction: rename it (and move it, if Manual File Mode is on).
        Returns False (and shows a message) if the fields aren't valid yet —
        callers should not close/advance in that case."""
        iid = self._correction_iid
        if not iid or iid not in self._iid_to_result:
            return True
        result = self._iid_to_result[iid]

        client  = self.corr_client_var.get().strip()
        subject = self.corr_subject_var.get().strip()
        if not client:
            messagebox.showwarning("Missing Client", "Please select a client from the list.")
            return False
        if not subject:
            messagebox.showwarning("Missing Subject", "Please enter a subject for the document.")
            return False

        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        current_name = result.final_name
        if not os.path.isfile(os.path.join(scandocs, current_name)):
            current_name = result.original_name
        if not os.path.isfile(os.path.join(scandocs, current_name)):
            # File's already gone (moved/renamed outside the app, or by a
            # prior correction) — nothing to rename. Don't block Close/Next
            # over it, just leave this row as-is.
            messagebox.showwarning(
                "File Not Found",
                f"Could not locate \"{current_name}\" on disk — "
                "skipping the rename for this file.")
            return True

        ext      = os.path.splitext(current_name)[1].lower()
        safe_sub = FileProcessor._safe_subject(subject) or "Document"
        new_name = f"{client} - {safe_sub}{ext}"

        if new_name != current_name:
            new_name = FileProcessor._resolve_collision(scandocs, new_name, current_name)
            try:
                os.rename(os.path.join(scandocs, current_name), os.path.join(scandocs, new_name))
            except Exception as e:
                messagebox.showwarning(
                    "Rename Failed",
                    f"Could not rename the file:\n{e}\n\nMoving on without renaming.")
                return True
        else:
            new_name = current_name

        result.final_name           = new_name
        result.client               = client
        result.description          = subject
        result.status               = "renamed"
        result.audit_corrected_name = new_name

        vals = list(self.results_tree.item(iid, "values"))
        vals[2] = new_name   # New Name
        vals[3] = "OK"       # Status
        vals[4] = client     # Client
        self.results_tree.item(iid, values=vals, tags=("renamed",))

        # Optional move, if Manual File Mode has a destination configured
        if self.s_file_mode_var.get() and self.s_file_mode_manual_var.get():
            dest = self.fo_dest_var.get().strip()
            if dest and os.path.isdir(dest):
                self._fo_do_move(result, dest, iid)

        return True

    def _corr_close(self):
        """Apply the correction (best-effort) and hide the panel — closing
        should never be blocked, even if the rename couldn't be applied."""
        self._corr_commit()
        self._hide_manual_correction()

    def _corr_next(self):
        """Apply the correction (best-effort), then open the next row."""
        self._corr_commit()
        children = self.results_tree.get_children()
        if not children or self._correction_iid not in children:
            self._hide_manual_correction()
            return
        idx = children.index(self._correction_iid) + 1
        if idx >= len(children):
            self._hide_manual_correction()
            return
        self._open_manual_correction(children[idx])

    # ── Processing ────────────────────────────────────────────

    def _start_processing(self):
        errors = self.config_mgr.validate()
        if errors:
            messagebox.showerror(
                "Configuration Error",
                "\n".join(errors) + "\n\nPlease fix these in Settings.",
            )
            self.notebook.select(self._settings_tab_frame)
            return

        # Clear old results
        self._hide_manual_correction()
        for row in self.results_tree.get_children():
            self.results_tree.delete(row)
        for btn in self._correction_buttons.values():
            btn.destroy()
        self._correction_buttons.clear()
        self._results.clear()
        self._iid_to_result.clear()
        self._total_files = 0
        self.progress_var.set(0)
        # Reset audit panel
        self.audit_file_label.config(
            text="Select a row above to audit it.", foreground="gray"
        )
        self._audit_updating = True
        self.audit_correct_var.set(False)
        self.audit_wrong_client_var.set(False)
        self.audit_bad_desc_var.set(False)
        self.audit_failed_client_var.set(False)
        self.audit_should_flag_var.set(False)
        self._audit_updating = False
        for w in (self.audit_open_btn, self.audit_prev_btn, self.audit_next_btn,
                  self.audit_correct_chk, self.audit_wrong_client_chk,
                  self.audit_bad_desc_chk, self.audit_failed_client_chk,
                  self.audit_should_flag_chk):
            w.config(state=tk.DISABLED)
        self.status_var.set("Starting…")
        self.btn_process.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        # Reset audit submit button for new run
        self.btn_submit_audit.config(state=tk.NORMAL, text="Submit Audit")
        self._processing_active = True

        t = threading.Thread(
            target=self.engine.run_batch,
            args=(self.config_mgr.config, self._queue),
            daemon=True,
        )
        t.start()
        self.after(100, self._poll_queue)

    def _stop_processing(self):
        self.engine.stop()
        self.status_var.set("Stopping after current file…")
        self.btn_stop.config(state=tk.DISABLED)

    def _poll_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                mtype = msg["type"]

                if mtype == "total":
                    self._total_files = msg["count"]
                    self.status_var.set(
                        f"Found {self._total_files} file(s). Processing…"
                    )

                elif mtype == "progress":
                    n = msg["current"]
                    pct = (n / self._total_files * 100) if self._total_files else 0
                    self.progress_var.set(pct)
                    self.status_var.set(
                        f"Processing {n} / {self._total_files}: {msg['filename']}"
                    )

                elif mtype == "result":
                    result = msg["result"]
                    # Suggest Location: pre-fill pending_dest if enabled
                    if (self.s_suggest_loc_var.get()
                            and self.s_file_mode_var.get()
                            and result.status in ("renamed", "needs_review")
                            and result.client
                            and not result.pending_dest):
                        _parent = self.s_suggest_parent_var.get().strip()
                        _desc = result.description
                        suggested = self._suggest_location(result.client, _desc, _parent)
                        if suggested:
                            result.pending_dest = suggested
                    self._add_result_row(result)

                elif mtype == "error":
                    messagebox.showerror("Error", msg["message"])
                    self._finish_processing()
                    return

                elif mtype == "stopped":
                    self.status_var.set("Stopped by user.")
                    self._finish_processing()
                    return

                elif mtype == "done":
                    self.progress_var.set(100)
                    self.status_var.set(
                        f"Done. {len(self._results)} file(s) processed."
                    )
                    self._finish_processing(auto_save=True)
                    self._refresh_review_tab()
                    self._refresh_unnamed_count()
                    return

        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _finish_processing(self, auto_save: bool = False):
        self._processing_active = False
        self.btn_process.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        # If the user finished auditing while processing was still running,
        # the audit-complete check was deferred — run it now.
        self._check_audit_complete()
        if auto_save and self._results and self.s_auto_save_var.get():
            folder = self.s_report_folder_var.get().strip() or DEFAULT_REPORTS_FOLDER
            try:
                path = self._write_report(folder)
                self.status_var.set(
                    self.status_var.get() + f"  |  Report saved: {os.path.basename(path)}"
                )
            except Exception as e:
                messagebox.showerror("Auto-Save Failed", f"Could not save report:\n{e}")

    # ── Report helpers ────────────────────────────────────────

    _REPORT_HEADERS = [
        "Original File", "New Name", "Status",
        "Client", "Description", "Confidence",
        "Extraction Method", "Renamed At", "Error",
        "Moved To",
        "Audit: Correct", "Audit: Wrong Client", "Audit: Bad Description",
        "Audit: Failed to Identify Client", "Audit: Should Have Flagged",
        "Audit: Corrected Name",
    ]

    def _results_as_rows(self) -> list:
        rows = []
        for r in self._results:
            rows.append([
                r.original_name, r.final_name, r.status,
                r.client, r.description, r.confidence,
                r.extraction_method, r.renamed_at or "", r.error_message or "",
                r.moved_to,
                "Yes" if r.audit_correct         else "",
                "Yes" if r.audit_wrong_client    else "",
                "Yes" if r.audit_bad_description else "",
                "Yes" if r.audit_failed_client   else "",
                "Yes" if r.audit_should_review   else "",
                r.audit_corrected_name,
            ])
        # Summary row
        total      = len(self._results)
        renamed    = sum(1 for r in self._results if r.status == "renamed")
        flagged    = sum(1 for r in self._results if r.status == "needs_review")
        moved      = sum(1 for r in self._results if r.moved_to)
        correct    = sum(1 for r in self._results if r.audit_correct)
        wrong_cl   = sum(1 for r in self._results if r.audit_wrong_client)
        bad_desc   = sum(1 for r in self._results if r.audit_bad_description)
        failed_cl  = sum(1 for r in self._results if r.audit_failed_client)
        sh_review  = sum(1 for r in self._results if r.audit_should_review)
        rows.append([])   # blank separator
        rows.append([
            "SUMMARY", "", "",
            f"Total: {total}", f"Renamed: {renamed}", f"Auto-flagged: {flagged}",
            "", "", "",
            f"Moved: {moved}",
            f"Confirmed correct: {correct}",
            f"Wrong client: {wrong_cl}",
            f"Bad description: {bad_desc}",
            f"Failed to identify client: {failed_cl}",
            f"Should have flagged: {sh_review}",
            "",
        ])
        return rows

    @staticmethod
    def _save_xlsx(path: str, headers: list, rows: list, results: list = None):
        """Write an Excel workbook with auto-fitted columns and a styled header row.
        Rows where 'Audit: Wrong Client' is 'Yes' are highlighted in red — these
        are the most critical errors and must be easy to spot.
        If `results` is provided, a Summary sheet is added."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Results"

        # Locate the "Audit: Wrong Client" column index (0-based in the data row)
        try:
            wrong_client_idx = headers.index("Audit: Wrong Client")
        except ValueError:
            wrong_client_idx = -1

        # Header row
        ws.append(headers)
        header_fill = PatternFill("solid", fgColor="1F497D")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 18

        # Fill styles
        light     = PatternFill("solid", fgColor="EEF2F7")
        red_fill  = PatternFill("solid", fgColor="FFCCCC")  # wrong-client flag
        red_font  = Font(bold=True, color="990000")

        for i, row in enumerate(rows, start=2):
            ws.append(row)
            # Detect wrong-client flag
            is_wrong_client = (
                wrong_client_idx >= 0
                and len(row) > wrong_client_idx
                and row[wrong_client_idx] == "Yes"
            )
            if is_wrong_client:
                for cell in ws[i]:
                    cell.fill = red_fill
                    cell.font = red_font
            elif i % 2 == 0:
                for cell in ws[i]:
                    cell.fill = light

        # Auto-fit column widths (cap between 12 and 60 chars wide)
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            letter = col[0].column_letter
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 60)

        ws.freeze_panes = "A2"  # keep header visible while scrolling

        # ── Summary sheet ─────────────────────────────────────
        if results is not None:
            ss = wb.create_sheet("Summary", 0)  # insert as first tab

            total       = len(results)
            skipped     = sum(1 for r in results if r.status == "skipped")
            errors      = sum(1 for r in results if r.status == "error")
            identified  = sum(1 for r in results if r.status == "renamed")
            needs_rev   = sum(1 for r in results if r.status == "needs_review")
            processed   = total - skipped - errors  # docs that went through AI

            # Audit counts (only processed files can be audited)
            _was_audited = lambda r: (r.audit_correct or r.audit_wrong_client
                                      or r.audit_bad_description or r.audit_failed_client
                                      or r.audit_should_review)
            audited       = sum(1 for r in results if _was_audited(r))
            audit_correct = sum(1 for r in results if r.audit_correct)
            audit_wrong   = sum(1 for r in results if r.audit_wrong_client)
            audit_bad_desc = sum(1 for r in results if r.audit_bad_description)
            audit_failed  = sum(1 for r in results if r.audit_failed_client)
            audit_should  = sum(1 for r in results if r.audit_should_review)

            # Identified-client subset audit
            id_results    = [r for r in results if r.status == "renamed"]
            id_audited    = sum(1 for r in id_results if _was_audited(r))
            id_correct    = sum(1 for r in id_results if r.audit_correct)

            def _pct(num, denom):
                return f"{num / denom * 100:.1f}%" if denom else "N/A"

            summary_data = [
                ("Documents Total",                     total),
                ("Documents Processed",                 processed),
                ("Documents Skipped",                   skipped),
                ("Documents With Errors",               errors),
                ("Documents Identified a Client",       identified),
                ("Documents Marked Needs Review",       needs_rev),
                ("", ""),
                ("Audit Completion (All Processed)",    _pct(audited, processed)),
                ("Audit Completion (Identified Only)",  _pct(id_audited, identified)),
                ("", ""),
                ("Success Rate (All Audited)",          _pct(audit_correct, audited)),
                ("Success Rate (Identified Only)",      _pct(id_correct, id_audited)),
                ("", ""),
                ("Audit Flags",                         ""),
                ("  Marked Correct",                    audit_correct),
                ("  Wrong Client",                      audit_wrong),
                ("  Bad Description",                   audit_bad_desc),
                ("  Failed to Identify Client",         audit_failed),
                ("  Should Have Been Flagged",          audit_should),
            ]

            # Style the summary sheet
            title_font = Font(bold=True, size=14, color="1F497D")
            label_font = Font(bold=True, size=11)
            value_font = Font(size=11)
            ss.append(["Audit Summary"])
            ss["A1"].font = title_font
            ss.append([])  # blank row

            for label, value in summary_data:
                ss.append([label, value])
            for row_cells in ss.iter_rows(min_row=3, max_row=ss.max_row, max_col=2):
                row_cells[0].font = label_font
                row_cells[1].font = value_font
                row_cells[1].alignment = Alignment(horizontal="right")

            ss.column_dimensions["A"].width = 42
            ss.column_dimensions["B"].width = 18

        wb.save(path)

    @staticmethod
    def _save_csv(path: str, headers: list, rows: list):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(headers)
            for row in rows:
                w.writerow(row)

    def _write_report(self, folder: str) -> str:
        """Write results to a timestamped report in `folder`. Returns the saved path."""
        os.makedirs(folder, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        headers = self._REPORT_HEADERS
        rows = self._results_as_rows()
        if _XLSX_AVAILABLE:
            path = os.path.join(folder, f"scandocs_report_{ts}.xlsx")
            self._save_xlsx(path, headers, rows, results=self._results)
        else:
            path = os.path.join(folder, f"scandocs_report_{ts}.csv")
            self._save_csv(path, headers, rows)
        return path


    def _add_result_row(self, result: ProcessResult):
        tag = result.status
        label = {
            "renamed":      "OK",
            "needs_review": "REVIEW",
            "skipped":      "Skipped",
            "error":        "ERROR",
        }.get(result.status, result.status)

        client_cell = result.client if result.client else (result.error_message or "")[:60]
        iid = self.results_tree.insert(
            "", tk.END,
            values=(
                "",   # audited checkmark — filled in by _on_audit_check
                result.original_name,
                result.final_name,
                label,
                client_cell,
                "",   # new_location — filled in by _fo_do_move
                "",   # correction — covered by a real overlaid button, see below
            ),
            tags=(tag,),
        )
        self._iid_to_result[iid] = result
        self._results.append(result)
        self._make_correction_button(iid)
        self.results_tree.yview_moveto(1.0)
        self._reposition_all_correction_buttons()

    def _make_correction_button(self, iid: str):
        """Create the blue 'Manual Correction' button overlaid on this row's
        correction cell. ttk.Treeview can't style an individual cell, so a
        real button is placed on top of it instead — see
        _position_correction_button for how it tracks the row."""
        btn = tk.Button(
            self.results_tree,
            text="✎  Manual Correction",
            command=lambda i=iid: self._open_manual_correction(i),
            bg="#1565c0", fg="white",
            activebackground="#0d47a1", activeforeground="white",
            disabledforeground="#cfd8dc",
            relief="solid", bd=1,
            font=(APP_FONT, 8, "bold"),
            cursor="hand2",
        )
        self._correction_buttons[iid] = btn
        self._position_correction_button(iid)

    def _position_correction_button(self, iid: str):
        """Move/show/hide the given row's correction button to track the
        treeview's current scroll position and column layout."""
        btn = self._correction_buttons.get(iid)
        if btn is None or not btn.winfo_exists():
            return
        bbox = self.results_tree.bbox(iid, "correction")
        if not bbox:
            btn.place_forget()
            return
        x, y, w, h = bbox
        btn.place(x=x + 2, y=y + 2, width=max(w - 4, 10), height=max(h - 4, 10))

    def _reposition_all_correction_buttons(self):
        for iid in list(self._correction_buttons.keys()):
            self._position_correction_button(iid)

    # ── Audit helpers ─────────────────────────────────────────

    @staticmethod
    def _is_audited(result: ProcessResult) -> bool:
        return (result.audit_correct or result.audit_wrong_client or
                result.audit_bad_description or result.audit_failed_client or
                result.audit_should_review)

    def _audit_mode_on(self) -> bool:
        return self.s_audit_mode_var.get()

    def _apply_extraction_method_ui(self):
        """Enable/disable the extraction-method dropdown based on whether the
        currently selected model is vision-capable. Also greys out the
        Max Vision Pages entry when OCR mode is selected."""
        if not hasattr(self, "s_method_combo"):
            return
        model = self.s_model_var.get().strip()
        vision_ok = model_supports_vision(model)

        if vision_ok:
            self.s_method_combo.configure(
                values=["Use OCR", "Use Vision Model"], state="readonly"
            )
            self._method_hint_lbl.configure(
                text=f"  '{model}' supports vision — both modes available."
            )
        else:
            # Force OCR and lock the dropdown
            if self.s_extraction_method_var.get() == "Use Vision Model":
                self.s_extraction_method_var.set("Use OCR")
            self.s_method_combo.configure(
                values=["Use OCR"], state="disabled"
            )
            if model:
                self._method_hint_lbl.configure(
                    text=f"  '{model}' is not on the vision allowlist "
                         f"(llama3.2-vision, gemma4). Vision mode disabled."
                )
            else:
                self._method_hint_lbl.configure(
                    text="  Select a model first. Vision mode requires "
                         "llama3.2-vision or gemma4."
                )

        # Grey out Max Vision Pages when not in vision mode
        in_vision = self.s_extraction_method_var.get() == "Use Vision Model"
        if hasattr(self, "s_vision_pages_entry"):
            self.s_vision_pages_entry.configure(
                state=("normal" if in_vision else "disabled")
            )

    def _apply_file_mode(self):
        """Show or hide File Mode panels based on the master toggle and per-tab sub-settings."""
        master_on = self.s_file_mode_var.get()
        auto_on   = master_on and self.s_file_mode_auto_var.get()
        manual_on = master_on and self.s_file_mode_manual_var.get()

        # Show / hide the per-tab sub-checkboxes
        if hasattr(self, "_file_mode_sub_frame"):
            if master_on:
                self._file_mode_sub_frame.grid()
            else:
                self._file_mode_sub_frame.grid_remove()

        # Suggest Location and Auto-commit are always visible but permanently disabled (Coming Soon)

        # Process tab — Move Files panel
        if auto_on:
            self.file_ops_panel.pack(fill=tk.X, padx=10, pady=(6, 10))
        else:
            self.file_ops_panel.pack_forget()

        # Manual Correction panel — destination row
        if hasattr(self, "_corr_dest_widgets"):
            for w in self._corr_dest_widgets:
                if manual_on:
                    w.grid()
                else:
                    w.grid_remove()

        # Legacy Manual Entry tab — destination row
        if hasattr(self, "_review_dest_widgets"):
            for w in self._review_dest_widgets:
                if manual_on:
                    w.grid()
                else:
                    w.grid_remove()

    def _browse_suggest_parent(self):
        folder = filedialog.askdirectory(
            title="Select Client Folders Parent Directory",
            initialdir=self.s_suggest_parent_var.get() or SCRIPT_DIR,
        )
        if folder:
            self.s_suggest_parent_var.set(os.path.normpath(folder))

    @staticmethod
    def _suggest_location(client: str, description: str, parent_folder: str) -> Optional[str]:
        """Find the single client folder in parent_folder that matches client, then pick
        the best subfolder based on the document description.
        Returns a path string, or None if no unambiguous match is found.
        """
        if not client or not parent_folder or not os.path.isdir(parent_folder):
            return None

        try:
            all_dirs = [
                d for d in os.listdir(parent_folder)
                if os.path.isdir(os.path.join(parent_folder, d))
            ]
        except Exception as e:
            logging.warning(f"Suggest location: could not list {parent_folder}: {e}")
            return None

        # Match the client name against folder names — look for the last name
        # (before the comma) plus partial first name to avoid false positives
        client_lower = client.lower()
        last_name = client.split(",")[0].strip().lower()
        first_part = client.split(",")[1].strip().lower().split()[0] if "," in client else ""

        matching_dirs = []
        for d in all_dirs:
            d_lower = d.lower()
            # Must contain the last name
            if last_name not in d_lower:
                continue
            # If we have a first name part, folder should contain it too
            if first_part and first_part not in d_lower:
                continue
            # Fuzzy fallback: ratio against the full client name
            ratio = difflib.SequenceMatcher(None, client_lower, d_lower).ratio()
            if ratio >= 0.7 or (last_name in d_lower and (not first_part or first_part in d_lower)):
                matching_dirs.append(d)

        if len(matching_dirs) != 1:
            # 0 matches → no folder found; 2+ matches → ambiguous, don't guess
            logging.debug(
                f"Suggest location: {len(matching_dirs)} matches for '{client}' "
                f"in {parent_folder} — skipping"
            )
            return None

        client_folder = os.path.join(parent_folder, matching_dirs[0])

        # Look for best subfolder using description keywords
        try:
            subfolders = [
                d for d in os.listdir(client_folder)
                if os.path.isdir(os.path.join(client_folder, d))
            ]
        except Exception:
            return client_folder

        if not subfolders:
            return client_folder

        desc_words = [w.lower() for w in re.split(r"\W+", description) if len(w) >= 3]
        if not desc_words:
            return client_folder

        best_sub = None
        best_score = 0
        for sub in subfolders:
            sub_lower = sub.lower()
            score = sum(1 for w in desc_words if w in sub_lower)
            if score > best_score:
                best_score = score
                best_sub = sub

        if best_sub and best_score >= 1:
            return os.path.join(client_folder, best_sub)
        return client_folder

    def _fo_browse_dest(self):
        folder = filedialog.askdirectory(
            title="Select Destination Folder",
            initialdir=self.fo_dest_var.get() or SCRIPT_DIR,
        )
        if folder:
            self.fo_dest_var.set(os.path.normpath(folder))
            # Persist immediately
            cfg = self.config_mgr.config
            cfg["processing"]["file_mode_destination"] = self.fo_dest_var.get()
            self.config_mgr.save(cfg)

    def _fo_resolve_dest(self) -> str | None:
        """Return the destination path if valid, else show an error and return None."""
        dest = self.fo_dest_var.get().strip()
        if not dest:
            messagebox.showwarning(
                "No Destination",
                "Please select a destination folder before moving files."
            )
            return None
        if not os.path.isdir(dest):
            create = messagebox.askyesno(
                "Folder Not Found",
                f"The destination folder does not exist:\n{dest}\n\nCreate it?",
            )
            if not create:
                return None
            try:
                os.makedirs(dest, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Could not create folder:\n{e}")
                return None
        return dest

    def _fo_do_move(self, result: ProcessResult, dest: str, iid: str) -> bool:
        """Move a single file to dest. Returns True on success."""
        if not os.path.isdir(dest):
            try:
                os.makedirs(dest, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Could not create destination folder:\n{e}")
                return False
        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        src = os.path.join(scandocs, result.final_name)
        if not os.path.isfile(src):
            messagebox.showwarning(
                "File Not Found",
                f"Could not locate:\n{result.final_name}\n\n"
                "It may have already been moved or renamed."
            )
            return False
        # Collision avoidance at destination
        dst_name = result.final_name
        base, ext = os.path.splitext(dst_name)
        counter = 1
        dst = os.path.join(dest, dst_name)
        while os.path.exists(dst):
            dst_name = f"{base} ({counter}){ext}"
            dst = os.path.join(dest, dst_name)
            counter += 1
        try:
            import shutil
            shutil.move(src, dst)
            result.moved_to = dst
            # Update the New Location column (index 5) in the treeview row
            vals = list(self.results_tree.item(iid, "values"))
            vals[5] = os.path.basename(dest)
            self.results_tree.item(iid, values=vals, tags=("moved",))
            return True
        except Exception as e:
            messagebox.showerror("Move Failed", f"Could not move file:\n{e}")
            return False

    def _fo_apply_to_selected(self):
        """Stage the current destination for all selected rows without moving them yet.
        The staged location is shown in the New Location column as '(pending)'."""
        dest = self.fo_dest_var.get().strip()
        if not dest:
            self.fo_status_var.set("Set a destination folder first.")
            return
        sel = self.results_tree.selection()
        if not sel:
            self.fo_status_var.set("Select one or more files first.")
            return
        folder_label = os.path.basename(dest) or dest
        audit_on = self._audit_mode_on()
        count = 0
        skipped_audit = 0
        for iid in sel:
            result = self._iid_to_result.get(iid)
            if result and result.status == "renamed" and not result.moved_to:
                # In audit mode a file must be marked Correct before it can be moved
                if audit_on and not result.audit_correct:
                    skipped_audit += 1
                    continue
                result.pending_dest = dest
                vals = list(self.results_tree.item(iid, "values"))
                vals[5] = f"{folder_label} (pending)"
                self.results_tree.item(iid, values=vals)
                count += 1
        if count:
            msg = f"{count} file{'s' if count != 1 else ''} staged — click Move Files to commit."
            if skipped_audit:
                msg += f"  ({skipped_audit} skipped — not marked Correct in audit.)"
            self.fo_status_var.set(msg)
        else:
            if skipped_audit:
                self.fo_status_var.set(
                    f"No files staged — {skipped_audit} file(s) must be marked Correct in audit first.")
            else:
                self.fo_status_var.set(
                    "No eligible files selected (files must be renamed and not yet moved).")

    def _fo_move_all(self):
        """Move all files that have been staged via Apply to Selected.
        In audit mode, only files marked Correct are eligible."""
        audit_on = self._audit_mode_on()
        moveable = [
            (iid, r) for iid, r in self._iid_to_result.items()
            if r.status == "renamed" and not r.moved_to and r.pending_dest
            and (not audit_on or r.audit_correct)
        ]
        if not moveable:
            self.fo_status_var.set("No files staged — select files and click Apply to Selected first.")
            return
        moved, failed = 0, 0
        for iid, result in moveable:
            if self._fo_do_move(result, result.pending_dest, iid):
                moved += 1
            else:
                failed += 1
        msg = f"Moved {moved} file{'s' if moved != 1 else ''}."
        if failed:
            msg += f"  {failed} failed."
        self.fo_status_var.set(msg)

    # ── Palette & rounding ────────────────────────────────────────────────────

    def _apply_default_styling(self):
        """Apply fixed default blue styling across all UI elements."""
        s = self.style
        primary = _APP_PRIMARY
        light   = _APP_LIGHT
        mid     = _APP_MID

        # Primary buttons
        s.configure("primary.TButton",
            background=primary, bordercolor=primary,
            darkcolor=primary,  lightcolor=primary,
            foreground="#ffffff",
        )
        s.map("primary.TButton",
            background=[("active !disabled", mid),
                        ("pressed !disabled", primary),
                        ("disabled", "#cccccc")],
            bordercolor=[("active !disabled", mid),
                         ("focus !disabled", mid)],
            darkcolor=[("active !disabled", mid),
                       ("pressed !disabled", primary)],
            lightcolor=[("active !disabled", mid),
                        ("pressed !disabled", primary)],
        )

        # Outline buttons
        s.configure("primary-outline.TButton",
            foreground=primary, bordercolor=primary,
        )
        s.map("primary-outline.TButton",
            background=[("active !disabled", light)],
            foreground=[("active !disabled", primary)],
        )

        # Progressbar
        s.configure("primary.Horizontal.TProgressbar",
            troughcolor="#e0e0e0", background=primary)
        s.configure("Horizontal.TProgressbar",
            troughcolor="#e0e0e0", background=primary)

        # Entry & Combobox focus highlight
        s.map("TEntry",
            bordercolor=[("focus !disabled", primary), ("hover !disabled", mid)],
            lightcolor=[("focus !disabled", light)],
            darkcolor=[("focus !disabled", light)],
        )
        s.map("TCombobox",
            bordercolor=[("focus !disabled", primary), ("hover !disabled", mid)],
            lightcolor=[("focus !disabled", light)],
        )

        # Notebook active tab
        s.map("TNotebook.Tab",
            background=[("selected", light), ("active", "#f8f9fa")],
            foreground=[("selected", primary)],
        )

        # LabelFrame title
        s.configure("TLabelframe.Label", foreground=primary)

        # Checkbutton / Radiobutton indicators
        s.map("TCheckbutton",  indicatorcolor=[("selected", primary)])
        s.map("Checkbutton",   indicatorcolor=[("selected", primary)])
        s.map("TRadiobutton",  indicatorcolor=[("selected", primary)])

        # Scrollbar thumb
        s.configure("Vertical.TScrollbar",   troughcolor="#f0f0f0", background=mid)
        s.configure("Horizontal.TScrollbar", troughcolor="#f0f0f0", background=mid)

    def _apply_round_styling(self):
        """Apply soft visual tweaks: padded entries/tabs and DWM rounded window corners."""
        s = self.style

        s.configure("TEntry",        padding=[8, 6])
        s.configure("TCombobox",     padding=[8, 6])
        s.configure("TNotebook.Tab", padding=[14, 6])
        s.configure("TButton",       padding=[10, 6])

        self._apply_default_styling()

        # Rounded window corners (Windows 11 DWM — silent no-op elsewhere)
        self.after(200, self._try_round_window)

    # ── Scroll helper ─────────────────────────────────────────

    @staticmethod
    def _bind_mousewheel(widget, scroll_fn):
        """Recursively bind mousewheel/trackpad scroll to widget and all descendants."""
        widget.bind("<MouseWheel>",  lambda e: scroll_fn(e), add="+")
        widget.bind("<Button-4>",    lambda e: scroll_fn(e), add="+")  # Linux scroll-up
        widget.bind("<Button-5>",    lambda e: scroll_fn(e), add="+")  # Linux scroll-down
        for child in widget.winfo_children():
            ScandocsApp._bind_mousewheel(child, scroll_fn)

    def _try_round_window(self):
        """Ask Windows 11 DWM to use rounded window corners."""
        try:
            import ctypes
            DWMWA_WINDOW_CORNER_PREFERENCE = 33
            DWMWCP_ROUND = 2
            hwnd = self.winfo_id()
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_WINDOW_CORNER_PREFERENCE,
                ctypes.byref(ctypes.c_int(DWMWCP_ROUND)),
                ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass  # Not Windows 11 or DWM unavailable — ignore

    def _apply_audit_mode(self):
        """Show or hide all audit UI based on the current Audit Mode setting."""
        on = self._audit_mode_on()
        # ✓ column in treeview
        if on:
            self.results_tree["displaycolumns"] = (
                "audited", "original", "new_name", "status", "client",
                "new_location", "correction")
        else:
            self.results_tree["displaycolumns"] = (
                "original", "new_name", "status", "client",
                "new_location", "correction")
        if hasattr(self, "_correction_buttons"):
            self.results_tree.after_idle(self._reposition_all_correction_buttons)
        # Audit panel
        if on:
            self.audit_panel.pack(fill=tk.X, padx=10, pady=(6, 0))
        else:
            self.audit_panel.pack_forget()
        # Top-right button: Submit Audit in audit mode, Open Report otherwise
        if on:
            self.btn_open_report.pack_forget()
            self.btn_submit_audit.pack(anchor="e")
        else:
            self.btn_submit_audit.pack_forget()
            self.btn_open_report.pack(anchor="e")

    def _check_audit_complete(self):
        """If all auditable rows are reviewed AND processing is done, prompt to submit."""
        if not self._audit_mode_on() or not self._results:
            return
        if getattr(self, "_processing_active", False):
            return  # Still processing — re-checked automatically when processing finishes
        auditable = [r for r in self._results if r.status in ("renamed", "needs_review")]
        if not auditable:
            return
        if all(self._is_audited(r) for r in auditable):
            folder = self.s_report_folder_var.get().strip() or DEFAULT_REPORTS_FOLDER
            try:
                self._write_report(folder)
            except Exception as e:
                messagebox.showerror("Report Error", f"Could not save report:\n{e}")
                return
            # Disable the submit button so it can't be double-submitted via auto-path
            self.btn_submit_audit.config(state=tk.DISABLED, text="Audit Saved")
            # Custom prompt — two explicit choices instead of a plain showinfo
            dlg = tk.Toplevel(self)
            dlg.title("Audit Complete")
            dlg.resizable(False, False)
            dlg.transient(self)
            dlg.grab_set()
            # Centre over the main window
            self.update_idletasks()
            x = self.winfo_x() + (self.winfo_width() - 380) // 2
            y = self.winfo_y() + (self.winfo_height() - 150) // 2
            dlg.geometry(f"380x150+{x}+{y}")
            ttk.Label(
                dlg,
                text="You finished the audit.\nWould you like to submit the audit now?",
                font=(APP_FONT, 11),
                anchor="center",
                justify="center",
            ).pack(pady=(24, 16))
            btn_row = ttk.Frame(dlg)
            btn_row.pack()
            def _do_submit():
                dlg.destroy()
                self.btn_submit_audit.config(state=tk.NORMAL, text="Submit Audit")
                self._submit_audit()
            def _make_changes():
                dlg.destroy()
                self.btn_submit_audit.config(state=tk.NORMAL, text="Submit Audit")
            ttk.Button(btn_row, text="Submit Audit", bootstyle="primary",
                       command=_do_submit).pack(side=tk.LEFT, padx=(0, 12))
            ttk.Button(btn_row, text="Make Changes", bootstyle="secondary-outline",
                       command=_make_changes).pack(side=tk.LEFT)

    def _on_close(self):
        """Warn if audit mode is on and items are still unreviewed."""
        if self._audit_mode_on() and self._results:
            auditable = [r for r in self._results
                         if r.status in ("renamed", "needs_review")]
            unreviewed = [r for r in auditable if not self._is_audited(r)]
            if unreviewed:
                n = len(unreviewed)
                answer = messagebox.askyesno(
                    "Audit Incomplete",
                    f"{n} file{'s' if n != 1 else ''} still "
                    f"{'have' if n != 1 else 'has'} not been audited.\n\n"
                    "You should complete the audit before closing so the report\n"
                    "includes accurate quality data.\n\n"
                    "Close anyway?",
                    icon="warning",
                )
                if not answer:
                    return
        self.destroy()

    def _sort_treeview(self, col: str):
        """Sort the results treeview by *col*, toggling A→Z / Z→A on repeated clicks."""
        col_index = {
            "audited": 0, "original": 1, "new_name": 2,
            "status": 3, "client": 4, "new_location": 5,
        }
        col_labels = {
            "audited": "✓", "original": "Original File", "new_name": "New Name",
            "status": "Status", "client": "Client",
            "new_location": "New Location",
        }
        sortable = {"original", "new_name", "status", "client"}

        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col = col
            self._sort_reverse = False

        idx = col_index[col]
        items = [(self.results_tree.item(iid, "values"), iid)
                 for iid in self.results_tree.get_children()]
        items.sort(key=lambda x: (x[0][idx] or "").lower(),
                   reverse=self._sort_reverse)
        for i, (_, iid) in enumerate(items):
            self.results_tree.move(iid, "", i)
        self._reposition_all_correction_buttons()

        # Update heading arrows
        arrow = " ↓" if self._sort_reverse else " ↑"
        for c, lbl in col_labels.items():
            if c not in sortable:
                continue
            self.results_tree.heading(
                c, text=lbl + (arrow if c == col else ""),
                command=lambda _c=c: self._sort_treeview(_c))

    def _on_result_select(self, _event=None):
        """Update the audit panel when the user selects a row."""
        sel = self.results_tree.selection()
        if not sel:
            return
        iid = sel[0]
        result = self._iid_to_result.get(iid)
        if not result:
            return

        # Update file label
        self.audit_file_label.config(
            text=result.final_name or result.original_name, foreground="#212529"
        )
        # Enable controls
        for w in (self.audit_open_btn, self.audit_prev_btn, self.audit_next_btn,
                  self.audit_correct_chk, self.audit_wrong_client_chk,
                  self.audit_bad_desc_chk, self.audit_failed_client_chk,
                  self.audit_should_flag_chk):
            w.config(state=tk.NORMAL)

        # Sync checkboxes without triggering callbacks
        self._audit_updating = True
        self.audit_correct_var.set(result.audit_correct)
        self.audit_wrong_client_var.set(result.audit_wrong_client)
        self.audit_bad_desc_var.set(result.audit_bad_description)
        self.audit_failed_client_var.set(result.audit_failed_client)
        self.audit_should_flag_var.set(result.audit_should_review)
        self._audit_updating = False

        # Sync the rename hint
        self._update_audit_rename_hint(result)

        # If the preview popup is open, refresh it with the newly selected file
        if self._file_popup is not None and self._file_popup.winfo_exists():
            scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
            path = os.path.join(scandocs, result.final_name)
            if not os.path.isfile(path):
                path = os.path.join(scandocs, result.original_name)
            if os.path.isfile(path):
                self._refresh_popup_content(path)


    def _audit_open_file(self):
        """Open the currently selected file in a popup viewer."""
        sel = self.results_tree.selection()
        if not sel:
            return
        result = self._iid_to_result.get(sel[0])
        if not result:
            return
        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        path = os.path.join(scandocs, result.final_name)
        if not os.path.isfile(path):
            path = os.path.join(scandocs, result.original_name)
        if os.path.isfile(path):
            self._open_file_popup(path)
        else:
            messagebox.showwarning("File Not Found",
                f"Could not locate the file:\n{result.final_name}")

    def _update_audit_rename_hint(self, result: ProcessResult):
        """Update the italic rename-hint label below the audit checkboxes.
        Wrong client name → client portion becomes A-NEEDS REVIEW.
        Bad description   → description portion becomes Scanned Document.
        Both flags can apply simultaneously."""
        any_bad = (result.audit_wrong_client or result.audit_bad_description
                   or result.audit_failed_client or result.audit_should_review)
        if not any_bad:
            self.audit_rename_hint_var.set("")
            return

        fname = result.final_name
        base, ext = os.path.splitext(fname)
        if " - " in base:
            orig_client, orig_desc = base.split(" - ", 1)
        else:
            orig_client, orig_desc = base, ""

        # Which part changes?
        client_bad = result.audit_wrong_client or result.audit_failed_client or result.audit_should_review
        new_client = "A-NEEDS REVIEW" if client_bad else orig_client
        new_desc   = "Scanned Document" if result.audit_bad_description else orig_desc

        proposed = f"{new_client} - {new_desc}{ext}" if new_desc else f"{new_client}{ext}"
        self.audit_rename_hint_var.set(
            f"This document will be renamed \"{proposed}\" after the audit is complete.")

    def _on_audit_check(self, flag: str):
        """Called when an audit checkbox is toggled."""
        if self._audit_updating:
            return
        sel = self.results_tree.selection()
        if not sel:
            return
        iid = sel[0]
        result = self._iid_to_result.get(iid)
        if not result:
            return

        if flag == "correct":
            result.audit_correct = self.audit_correct_var.get()
            # Correct is mutually exclusive with all problem flags — clear them
            if result.audit_correct:
                result.audit_wrong_client    = False
                result.audit_bad_description = False
                result.audit_failed_client   = False
                result.audit_should_review   = False
                self._audit_updating = True
                self.audit_wrong_client_var.set(False)
                self.audit_bad_desc_var.set(False)
                self.audit_failed_client_var.set(False)
                self.audit_should_flag_var.set(False)
                self._audit_updating = False
        elif flag == "wrong_client":
            result.audit_wrong_client = self.audit_wrong_client_var.get()
        elif flag == "bad_description":
            result.audit_bad_description = self.audit_bad_desc_var.get()
        elif flag == "failed_client":
            result.audit_failed_client = self.audit_failed_client_var.get()
        elif flag == "should_review":
            result.audit_should_review = self.audit_should_flag_var.get()

        # Orange if any problem flag set; keep green if marked correct; else original
        any_flagged = (result.audit_wrong_client or result.audit_bad_description
                       or result.audit_failed_client or result.audit_should_review)
        if any_flagged:
            tags = ("audited",)
        elif result.audit_correct:
            tags = ("renamed",)
        else:
            tags = (result.status,)
        self.results_tree.item(iid, tags=tags)

        # Update the ✓ column
        is_audited = self._is_audited(result)
        vals = list(self.results_tree.item(iid, "values"))
        vals[0] = "✓" if is_audited else ""
        self.results_tree.item(iid, values=vals)

        # Check if all auditable rows are now done
        self._check_audit_complete()

        # Update the rename hint below the checkboxes
        self._update_audit_rename_hint(result)

    def _show_correction_dialog(self, iid: str, result: ProcessResult):
        """Modal dialog asking the employee what the file should be named."""
        dialog = tk.Toplevel(self)
        dialog.title("Correct File Name")
        dialog.resizable(False, False)
        dialog.grab_set()   # modal
        dialog.focus_set()

        # Centre over the main window
        self.update_idletasks()
        x = self.winfo_x() + self.winfo_width()  // 2 - 260
        y = self.winfo_y() + self.winfo_height() // 2 - 80
        dialog.geometry(f"520x160+{x}+{y}")

        ttk.Label(dialog, text="What should this file be named?",
                  font=(APP_FONT, 10, "bold")).pack(anchor="w", padx=16, pady=(16, 4))
        ttk.Label(dialog, text="Include the file extension  (e.g. GARCIA, Maria - Invoice.pdf)",
                  foreground="gray", font=(APP_FONT, 8)).pack(anchor="w", padx=16)

        entry_var = tk.StringVar(value=result.audit_corrected_name or result.final_name)
        entry = ttk.Entry(dialog, textvariable=entry_var, width=62)
        entry.pack(padx=16, pady=(6, 0), fill=tk.X)
        entry.select_range(0, tk.END)
        entry.focus_set()

        def _submit():
            corrected = entry_var.get().strip()
            if corrected:
                result.audit_corrected_name = corrected
                # Update the New Name cell in the treeview to show the correction
                vals = list(self.results_tree.item(iid, "values"))
                vals[1] = f"{corrected}  ✎"
                self.results_tree.item(iid, values=vals)
            dialog.destroy()

        def _cancel():
            # Uncheck whichever box triggered this if no correction was previously saved
            if not result.audit_corrected_name:
                self._audit_updating = True
                if result.audit_wrong_client and not self.audit_bad_desc_var.get():
                    self.audit_wrong_client_var.set(False)
                    result.audit_wrong_client = False
                elif result.audit_bad_description:
                    self.audit_bad_desc_var.set(False)
                    result.audit_bad_description = False
                self._audit_updating = False
                # Re-evaluate row colour
                any_flagged = (result.audit_wrong_client or result.audit_bad_description
                               or result.audit_should_review)
                tags = ("audited",) if any_flagged else (result.status,)
                self.results_tree.item(iid, tags=tags)
            dialog.destroy()

        btn_row = ttk.Frame(dialog)
        btn_row.pack(pady=(10, 0))
        ttk.Button(btn_row, text="Submit", bootstyle="primary",
                   command=_submit).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="Cancel", bootstyle="dark-outline",
                   command=_cancel).pack(side=tk.LEFT, padx=6)

        dialog.bind("<Return>", lambda e: _submit())
        dialog.bind("<Escape>", lambda e: _cancel())
        dialog.wait_window()

    # ── Audit navigation ──────────────────────────────────────

    def _audit_prev(self, _event=None):
        """Move selection to the previous row in the results tree."""
        children = self.results_tree.get_children()
        sel = self.results_tree.selection()
        if not children:
            return
        if not sel:
            self.results_tree.selection_set(children[-1])
            self.results_tree.see(children[-1])
        else:
            idx = children.index(sel[0])
            if idx > 0:
                self.results_tree.selection_set(children[idx - 1])
                self.results_tree.see(children[idx - 1])

    def _audit_next(self, _event=None):
        """Move selection to the next row in the results tree."""
        children = self.results_tree.get_children()
        sel = self.results_tree.selection()
        if not children:
            return
        if not sel:
            self.results_tree.selection_set(children[0])
            self.results_tree.see(children[0])
        else:
            idx = children.index(sel[0])
            if idx < len(children) - 1:
                self.results_tree.selection_set(children[idx + 1])
                self.results_tree.see(children[idx + 1])

    # ── File popup viewer ─────────────────────────────────────

    # ── Document popup viewer ─────────────────────────────────

    def _render_doc_image(self, path: str):
        """Render up to the first three PDF pages (or a JPEG) into a single PIL image.
        Returns None on any error so callers can fall back gracefully."""
        ext = os.path.splitext(path)[1].lower()
        try:
            import io
            if ext == ".pdf":
                doc = fitz.open(path)
                mat = fitz.Matrix(1.5, 1.5)
                n = min(doc.page_count, 3)
                imgs = [
                    PILImage.open(io.BytesIO(doc[i].get_pixmap(matrix=mat).tobytes("png")))
                    for i in range(n)
                ]
                doc.close()
                if len(imgs) == 1:
                    return imgs[0]
                # Stack pages vertically with a light-grey gap
                gap = 8
                w = max(p.width  for p in imgs)
                h = sum(p.height for p in imgs) + gap
                combined = PILImage.new("RGB", (w, h), color=(200, 200, 200))
                y = 0
                for p in imgs:
                    combined.paste(p, (0, y))
                    y += p.height + gap
                return combined
            elif ext in (".jpg", ".jpeg"):
                return PILImage.open(path)
        except Exception as e:
            logging.warning(f"Could not render {path}: {e}")
        return None

    def _open_file_popup(self, path: str):
        """Open the preview popup for *path*.
        If the popup is already visible, refresh its content instead of opening a second window."""
        if fitz is None or PILImage is None or PILImageTk is None:
            _open_file(path)
            return

        img = self._render_doc_image(path)
        if img is None:
            _open_file(path)
            return

        # Reuse existing window if it is still open
        if self._file_popup is not None and self._file_popup.winfo_exists():
            self._refresh_popup_content(path, img)
            return

        # ── Create new popup ──────────────────────────────────
        popup = tk.Toplevel(self)
        popup.title(os.path.basename(path))
        popup.resizable(True, True)

        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        ui_cfg = self.config_mgr.config.get("ui", {})
        saved_w = ui_cfg.get("preview_popup_width")  or 0
        saved_h = ui_cfg.get("preview_popup_height") or 0
        if saved_w > 0 and saved_h > 0:
            width, height = saved_w, saved_h
        else:
            # Default: large, centered popup rather than a small side panel
            width  = int(screen_w * 0.8)
            height = int(screen_h * 0.85)
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        popup.geometry(f"{width}x{height}+{x}+{y}")

        def _save_popup_size(_event=None):
            try:
                w = popup.winfo_width()
                h = popup.winfo_height()
            except tk.TclError:
                return
            if w <= 1 or h <= 1:
                return
            self.config_mgr.config["ui"]["preview_popup_width"]  = w
            self.config_mgr.config["ui"]["preview_popup_height"] = h
            self.config_mgr.save()

        self._popup_resize_after_id = None

        def _on_popup_configure(event):
            if event.widget is not popup:
                return
            if self._popup_resize_after_id is not None:
                popup.after_cancel(self._popup_resize_after_id)
            self._popup_resize_after_id = popup.after(500, _save_popup_size)

        popup.bind("<Configure>", _on_popup_configure)

        canvas = tk.Canvas(popup, bg="white")
        vsb = ttk.Scrollbar(popup, orient="vertical",   command=canvas.yview)
        hsb = ttk.Scrollbar(popup, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT,  fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(fill=tk.BOTH, expand=True)

        photo = PILImageTk.PhotoImage(img)
        canvas.create_image(0, 0, anchor="nw", image=photo)
        canvas.configure(scrollregion=(0, 0, img.width, img.height))
        canvas.image = photo

        canvas.bind("<MouseWheel>",
                    lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        self._file_popup        = popup
        self._file_popup_canvas = canvas

        def _on_popup_close():
            if self._popup_resize_after_id is not None:
                popup.after_cancel(self._popup_resize_after_id)
                self._popup_resize_after_id = None
            _save_popup_size()
            self._file_popup        = None
            self._file_popup_canvas = None
            popup.destroy()

        popup.protocol("WM_DELETE_WINDOW", _on_popup_close)

        # Force keyboard focus onto the popup itself — otherwise, if it was
        # opened while an Entry field had focus (e.g. from the Manual
        # Correction panel), spacebar would keep typing into that field
        # instead of reaching the popup's <space>-to-close binding below.
        popup.lift()
        popup.focus_force()

        # Spacebar while the popup is focused closes it, mirroring the
        # tree/listbox toggle behaviour (the popup steals keyboard focus,
        # so the tree's own <space> binding never fires while it's open).
        popup.bind("<space>", lambda e: _on_popup_close())

        # Arrow keys while the popup is focused still navigate the active list
        # (results tree on the Process tab, or the legacy Manual Entry list).
        popup.bind("<Left>",  lambda e: self._popup_prev())
        popup.bind("<Right>", lambda e: self._popup_next())
        popup.bind("<Up>",    lambda e: self._popup_prev())
        popup.bind("<Down>",  lambda e: self._popup_next())

    def _refresh_popup_content(self, path: str, img=None):
        """Swap the canvas image in the already-open popup for a new file."""
        if self._file_popup is None or not self._file_popup.winfo_exists():
            return
        if img is None:
            img = self._render_doc_image(path)
        if img is None:
            return
        canvas = self._file_popup_canvas
        photo  = PILImageTk.PhotoImage(img)
        canvas.delete("all")
        canvas.create_image(0, 0, anchor="nw", image=photo)
        canvas.configure(scrollregion=(0, 0, img.width, img.height))
        canvas.image = photo          # prevent GC
        canvas.yview_moveto(0)        # scroll back to top for the new document
        self._file_popup.title(os.path.basename(path))

    def _toggle_file_popup(self, path: str):
        """Open the popup if it is closed; close it if it is already open."""
        if self._file_popup is not None and self._file_popup.winfo_exists():
            self._file_popup.destroy()
            self._file_popup        = None
            self._file_popup_canvas = None
        else:
            self._open_file_popup(path)

    def _popup_prev(self):
        """Arrow-key navigation while the preview popup is focused — routes to
        whichever list is currently on screen (Process tab or the legacy
        Manual Entry tab)."""
        current = self.notebook.select()
        if current == str(self._process_tab_frame):
            self._audit_prev()
        elif current == str(self._review_tab_frame):
            self._review_prev()

    def _popup_next(self):
        current = self.notebook.select()
        if current == str(self._process_tab_frame):
            self._audit_next()
        elif current == str(self._review_tab_frame):
            self._review_next()

    def _on_tree_return(self, _event=None):
        """Spacebar on the Auto-Process results tree: toggle the file viewer."""
        sel = self.results_tree.selection()
        if not sel:
            return "break"
        result = self._iid_to_result.get(sel[0])
        if not result:
            return "break"
        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        path = os.path.join(scandocs, result.final_name)
        if not os.path.isfile(path):
            path = os.path.join(scandocs, result.original_name)
        if os.path.isfile(path):
            self._toggle_file_popup(path)
        return "break"   # prevent treeview default Enter behaviour

    def _on_tree_double_click(self, _event=None):
        """Double-click on the Auto-Process results tree: always open the file viewer."""
        sel = self.results_tree.selection()
        if not sel:
            return "break"
        result = self._iid_to_result.get(sel[0])
        if not result:
            return "break"
        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        path = os.path.join(scandocs, result.final_name)
        if not os.path.isfile(path):
            path = os.path.join(scandocs, result.original_name)
        if os.path.isfile(path):
            self._open_file_popup(path)
        return "break"

    # ── Submit Audit ──────────────────────────────────────────

    def _submit_audit(self):
        """Apply audit-flagged renames and save the report."""
        if not self._results:
            messagebox.showinfo("No Results", "No results to submit. Run processing first.")
            return
        scandocs = self.config_mgr.config["paths"]["scandocs_folder"]
        errors = []

        for result in self._results:
            if result.status not in ("renamed", "needs_review"):
                continue
            if not (result.audit_wrong_client or result.audit_bad_description):
                continue

            current_name = result.final_name
            base, ext = os.path.splitext(current_name)

            # Determine new client segment
            m = re.match(r"^(.+?) - (.+)$", base)
            old_client = m.group(1) if m else base
            old_desc   = m.group(2) if m else ""

            new_client = "A-NEEDS REVIEW" if result.audit_wrong_client else old_client
            new_desc   = "Scanned Document" if result.audit_bad_description else old_desc

            new_name = f"{new_client} - {new_desc}{ext}" if old_desc else f"{new_client}{ext}"

            if new_name == current_name:
                continue

            src = os.path.join(scandocs, current_name)
            if os.path.isfile(src):
                resolved = FileProcessor._resolve_collision(scandocs, new_name, current_name)
                try:
                    os.rename(src, os.path.join(scandocs, resolved))
                    result.final_name = resolved
                    result.audit_corrected_name = resolved
                    # Refresh the treeview row
                    for iid, r in self._iid_to_result.items():
                        if r is result:
                            vals = list(self.results_tree.item(iid, "values"))
                            vals[2] = resolved
                            self.results_tree.item(iid, values=vals)
                            break
                except Exception as e:
                    errors.append(f"{current_name}: {e}")

        # Build audit folder: ScandocsAudit_YYYY-MM-DD_NN
        import shutil
        report_root = self.s_report_folder_var.get().strip() or DEFAULT_REPORTS_FOLDER
        os.makedirs(report_root, exist_ok=True)
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        audit_num = 1
        while True:
            audit_folder_name = f"ScandocsAudit_{today_str}_{audit_num:02d}"
            audit_folder = os.path.join(report_root, audit_folder_name)
            if not os.path.exists(audit_folder):
                break
            audit_num += 1
        os.makedirs(audit_folder)

        # Save report into the audit folder
        try:
            path = self._write_report(audit_folder)
        except Exception as e:
            messagebox.showerror("Report Error", f"Could not save report:\n{e}")
            return

        # Copy files marked as Wrong Client into the audit folder
        wrong_client_copies = []
        for result in self._results:
            if not result.audit_wrong_client:
                continue
            src_path = os.path.join(scandocs, result.final_name)
            if not os.path.isfile(src_path):
                src_path = os.path.join(scandocs, result.original_name)
            if os.path.isfile(src_path):
                try:
                    dest = os.path.join(audit_folder, os.path.basename(src_path))
                    shutil.copy2(src_path, dest)
                    wrong_client_copies.append(os.path.basename(src_path))
                except Exception as e:
                    errors.append(f"Copy {os.path.basename(src_path)}: {e}")

        # Grey out the button so it can't be submitted twice
        self.btn_submit_audit.config(state=tk.DISABLED, text="Audit Submitted ✓")

        msg = f"Audit submitted.\nReport saved to:\n{audit_folder_name}/"
        if wrong_client_copies:
            msg += f"\n\n{len(wrong_client_copies)} Wrong Client file(s) copied to audit folder."
        if errors:
            msg += f"\n\nWarnings ({len(errors)}):\n" + "\n".join(errors[:5])
        messagebox.showinfo("Audit Submitted", msg)
        self._refresh_review_tab()

    # ── Client combo filter ───────────────────────────────────

    def _filter_client_combo(self, _event=None):
        """Filter the visible client list as the user types in the entry field."""
        typed = self.corr_client_var.get().lower()
        self.corr_client_listbox.delete(0, tk.END)
        for name in self._all_clients:
            if not typed or typed in name.lower():
                self.corr_client_listbox.insert(tk.END, name)

    def _on_client_listbox_select(self, _event=None):
        """Clicking a name in the client list copies it to the entry field."""
        sel = self.corr_client_listbox.curselection()
        if sel:
            self.corr_client_var.set(self.corr_client_listbox.get(sel[0]))

    def _open_report(self):
        folder = self.s_report_folder_var.get().strip() or DEFAULT_REPORTS_FOLDER
        if not os.path.isdir(folder):
            messagebox.showinfo("No Reports", f"Reports folder not found:\n{folder}")
            return
        reports = sorted(
            [f for f in os.listdir(folder)
             if f.lower().endswith(".xlsx") or f.lower().endswith(".csv")],
            reverse=True,
        )
        if not reports:
            messagebox.showinfo(
                "No Reports",
                f"No reports found in:\n{folder}\n\nProcess some documents first.",
            )
            return
        os.startfile(os.path.join(folder, reports[0]))

    def _export_csv(self):
        if not self._results:
            messagebox.showinfo("No Data", "No results to export. Run processing first.")
            return
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        init_dir = self.s_report_folder_var.get().strip() or DEFAULT_REPORTS_FOLDER
        if _XLSX_AVAILABLE:
            default_ext = ".xlsx"
            filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
            default_name = f"scandocs_report_{ts}.xlsx"
        else:
            default_ext = ".csv"
            filetypes = [("CSV files", "*.csv")]
            default_name = f"scandocs_report_{ts}.csv"
        path = filedialog.asksaveasfilename(
            title="Save Report",
            initialdir=init_dir,
            initialfile=default_name,
            defaultextension=default_ext,
            filetypes=filetypes,
        )
        if not path:
            return
        try:
            headers = self._REPORT_HEADERS
            rows = self._results_as_rows()
            if path.lower().endswith(".xlsx") and _XLSX_AVAILABLE:
                self._save_xlsx(path, headers, rows)
            else:
                self._save_csv(path, headers, rows)
            messagebox.showinfo("Exported", f"Report saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    # ── First-run wizard ──────────────────────────────────────

    def _check_first_run(self):
        if not self.config_mgr.config["paths"]["scandocs_folder"]:
            self._run_first_run_wizard()

    def _run_first_run_wizard(self):
        messagebox.showinfo(
            "Welcome to Speedy Scandocs",
            "Let's get you set up.\n\nFirst, select your Scandocs folder "
            "(where the scanner drops files).",
        )
        scandocs = filedialog.askdirectory(title="Select Scandocs Folder")
        if scandocs:
            p = os.path.normpath(scandocs)
            self.config_mgr.config["paths"]["scandocs_folder"] = p
            self.s_scandocs_var.set(p)

        messagebox.showinfo(
            "Client List File",
            "Now select your client_list.txt file, or choose a location to create one.\n\n"
            "This file stores client names in LAST, First format.",
        )
        client_file = filedialog.asksaveasfilename(
            title="Select or Create Client List",
            initialfile="client_list.txt",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
        )
        if client_file:
            p = os.path.normpath(client_file)
            self.config_mgr.config["paths"]["client_list_file"] = p
            self.s_client_list_var.set(p)
            if not os.path.exists(client_file):
                open(client_file, "w").close()

        self.config_mgr.save()
        messagebox.showinfo(
            "Setup Complete",
            "Setup complete!\n\n"
            "Next steps:\n"
            "1. Go to the Client List tab → add your clients → Save.\n"
            "2. Go to Settings → verify API URLs → Test Connection.\n"
            "3. Return to Process → click Auto-Process Documents.",
        )
        self.notebook.select(self._settings_tab_frame)

    # ── Auto-update ───────────────────────────────────────────

    def _maybe_check_for_updates_async(self):
        """Called on startup. Skips the check if disabled or if <24h since
        the last attempt. Runs the actual network call on a daemon thread
        so we never block the UI."""
        if IS_TEST_BUILD:
            logging.info("Auto-update check skipped: test build.")
            return
        cfg = self.config_mgr.config.get("updates", {})
        if not cfg.get("check_on_startup", True):
            return
        last = cfg.get("last_check_iso") or ""
        if last:
            try:
                last_dt = datetime.datetime.fromisoformat(last)
                elapsed = (datetime.datetime.now() - last_dt).total_seconds()
                if elapsed < UPDATE_CHECK_INTERVAL_SEC:
                    return
            except ValueError:
                pass
        self._check_for_updates_async(silent=True)

    def _check_for_updates_async(self, silent: bool = False):
        """Kick off the release lookup on a background thread.
        silent=True: no popup if we're up-to-date or offline.
        silent=False (manual "Check now"): always inform the user."""
        if IS_TEST_BUILD:
            logging.info("Auto-update check skipped: test build.")
            if not silent:
                messagebox.showinfo(
                    "Check for Updates",
                    "Auto-update is disabled in test builds.")
            return
        t = threading.Thread(
            target=self._do_update_check, args=(silent,), daemon=True,
        )
        t.start()

    def _do_update_check(self, silent: bool):
        release = fetch_latest_release()
        if not release:
            # Don't stamp last_check_iso on failure — otherwise a one-time
            # network blip (or a previously private repo returning 404) locks
            # auto-update out for 24h. The check only runs once per launch,
            # so retry-on-next-launch is fine.
            if not silent:
                self.after(0, lambda: messagebox.showwarning(
                    "Check for Updates",
                    "Couldn't reach the update server.\n"
                    "Check your internet connection and try again."))
            return

        self.config_mgr.config["updates"]["last_check_iso"] = \
            datetime.datetime.now().isoformat(timespec="seconds")
        try:
            self.config_mgr.save()
        except Exception:
            pass

        tag = release.get("tag_name") or ""
        latest = _parse_version(tag)
        current = _parse_version(APP_VERSION)
        if latest <= current:
            if not silent:
                self.after(0, lambda: messagebox.showinfo(
                    "Check for Updates",
                    f"You're up to date.\n\nInstalled version: {APP_VERSION}"))
            return

        asset = _pick_release_asset(release)
        if not asset:
            if not silent:
                self.after(0, lambda: messagebox.showinfo(
                    "Update Available",
                    f"Version {tag} is available, but no installer was found "
                    "for this platform. Visit the Releases page to download "
                    "it manually."))
            return

        # Respect a "skip this version" choice from a previous prompt.
        skip = self.config_mgr.config["updates"].get("skip_version", "")
        if silent and skip == tag:
            return

        notes = (release.get("body") or "").strip()
        self.after(0, lambda: self._prompt_update(tag, asset, notes))

    def _prompt_update(self, tag: str, asset: dict, notes: str):
        """Ask the user whether to install the newer version now."""
        short_notes = notes if len(notes) < 600 else notes[:600].rstrip() + "…"
        message = (
            f"A new version of Speedy Scandocs is available.\n\n"
            f"Installed: {APP_VERSION}\n"
            f"Available: {tag.lstrip('vV')}\n\n"
        )
        if short_notes:
            message += f"What's new:\n{short_notes}\n\n"
        message += "Install now?\n\n(The app will close and the installer will open.)"

        # yesnocancel: Yes = install, No = remind later, Cancel = skip this version
        resp = messagebox.askyesnocancel("Update Available", message)
        if resp is None:
            self.config_mgr.config["updates"]["skip_version"] = tag
            self.config_mgr.save()
            return
        if resp is False:
            return
        self._download_and_install(asset)

    def _download_and_install(self, asset: dict):
        """Show a progress dialog, stream the installer to a temp file,
        then launch it and exit the app."""
        import tempfile
        url = asset.get("browser_download_url")
        name = asset.get("name") or "SpeedyScandocsInstaller"
        if not url:
            return
        dest = os.path.join(tempfile.gettempdir(), name)

        # Progress dialog
        dlg = tk.Toplevel(self)
        dlg.title("Downloading Update")
        dlg.geometry("420x140")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        ttk.Label(dlg, text=f"Downloading {name}…", padding=(15, 12, 15, 0)).pack(anchor="w")
        pbar = ttk.Progressbar(dlg, mode="determinate", maximum=100, length=390)
        pbar.pack(padx=15, pady=10)
        status_var = tk.StringVar(value="Starting…")
        ttk.Label(dlg, textvariable=status_var, padding=(15, 0)).pack(anchor="w")

        cancelled = {"v": False}
        def _cancel():
            cancelled["v"] = True
            dlg.destroy()
        cancel_btn = ttk.Button(dlg, text="Cancel", command=_cancel,
                                bootstyle="secondary-outline")
        cancel_btn.pack(pady=6)

        def _progress(done, total):
            pct = int(done * 100 / total) if total else 0
            mb_done = done / (1024 * 1024)
            mb_total = total / (1024 * 1024) if total else 0
            self.after(0, lambda: (
                pbar.configure(value=pct),
                status_var.set(
                    f"{mb_done:.1f} MB / {mb_total:.1f} MB ({pct}%)"
                    if total else f"{mb_done:.1f} MB"
                ),
            ))

        def _worker():
            ok = download_file(url, dest, progress_cb=_progress,
                               cancel_flag=lambda: cancelled["v"])
            self.after(0, lambda: self._on_download_done(dlg, ok, dest, cancelled["v"]))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_download_done(self, dlg, ok: bool, path: str, was_cancelled: bool):
        try:
            dlg.destroy()
        except Exception:
            pass
        if was_cancelled:
            return
        if not ok:
            messagebox.showerror(
                "Update Failed",
                "The update could not be downloaded. Please try again later "
                "or download the installer manually from the Releases page.",
            )
            return
        try:
            if sys.platform == "win32":
                # Inno Setup installer handles UAC + overwrite-in-place itself.
                os.startfile(path)
            elif sys.platform == "darwin":
                # Opens the DMG in Finder; user drags the new .app to
                # Applications. Not seamless, but the standard Mac pattern.
                subprocess.Popen(["open", path])
            else:
                _open_file(path)
        except Exception as e:
            messagebox.showerror(
                "Update Failed",
                f"Could not launch the installer:\n{e}\n\n"
                f"Installer location:\n{path}",
            )
            return
        # Give the installer a moment to spawn before we exit.
        self.after(600, self._on_close)


# ─────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────

def _missing_dep_error(package: str):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing Dependency",
        f"{package} is not installed.\n\n"
        "Please run:\n    pip install -r requirements.txt\n\n"
        "Then restart Speedy Scandocs.",
    )
    sys.exit(1)


if __name__ == "__main__":
    logging.basicConfig(
        filename=LOG_PATH,
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    if fitz is None:
        _missing_dep_error("PyMuPDF")
    if requests is None:
        _missing_dep_error("requests")

    _load_bundled_fonts()

    app = ScandocsApp()
    app.mainloop()
