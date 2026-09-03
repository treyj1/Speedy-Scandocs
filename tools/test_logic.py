#!/usr/bin/env python3
"""
Speedy Scandocs — pure-logic regression tests (DEVELOPER TOOL ONLY)
===================================================================

Never bundled into the app. Complements tools/benchmark.py: that one
measures *naming accuracy* against real documents and needs a folder of
real client files, so it can't run unattended. This one tests the
deterministic logic around it — filename templates, collision handling,
the learning store's guards, config merging, batch orchestration — with no
documents, no network, and no model.

    python3 tools/test_logic.py

Exit code is 0 when everything passes, 1 otherwise, so it can gate a build.

Most cases here are pinned to a specific bug that reached a release
candidate; each is labelled with what it caught. Please add one whenever
you fix something in this layer — several of these were bugs that unit
tests this cheap would have caught before they ever shipped.

The app imports tkinter/ttkbootstrap at module scope. Those are present on
a normal dev machine, but this file installs lightweight stubs when they
aren't, so the logic can still be tested on a headless box or in CI.
"""

import os
import sys
import json
import types
import queue
import shutil
import tempfile
import importlib.util

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP_FILE = os.path.join(REPO_ROOT, "scandocs_tool.py")


def _install_gui_stubs_if_needed():
    """Let the module import on a machine with no tkinter/ttkbootstrap.
    Only the names touched at import time need to exist — nothing in this
    file exercises a widget."""
    try:
        import tkinter  # noqa: F401
        import ttkbootstrap  # noqa: F401
        return False
    except ImportError:
        pass

    class _Any:
        """Stands in for any widget class or constant. Must be a real class,
        not an instance — the app subclasses ttkbootstrap.Window."""

        def __init__(self, *a, **k):
            pass

        def __getattr__(self, name):
            return _Any()

        def __call__(self, *a, **k):
            return _Any()

    def _mk(name):
        mod = types.ModuleType(name)
        # Return the class itself so `class ScandocsApp(ttk.Window)` works.
        mod.__getattr__ = lambda n: _Any
        sys.modules[name] = mod
        return mod

    for name in ("tkinter", "tkinter.messagebox", "tkinter.filedialog",
                 "tkinter.font", "tkinter.ttk", "ttkbootstrap",
                 "ttkbootstrap.constants", "ttkbootstrap.scrolled"):
        _mk(name)
    return True


def _load_app():
    spec = importlib.util.spec_from_file_location("scandocs_tool_under_test", APP_FILE)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


STUBBED = _install_gui_stubs_if_needed()
sd = _load_app()

_PASSED = 0
_FAILED = []


def check(label, got, want):
    global _PASSED
    if got == want:
        _PASSED += 1
    else:
        _FAILED.append(f"{label}\n      got:  {got!r}\n      want: {want!r}")


def section(title):
    print(f"\n── {title}")


# ─────────────────────────────────────────────────────────────
# NameTemplate
# ─────────────────────────────────────────────────────────────

def test_name_template():
    section("NameTemplate")
    NT = sd.NameTemplate
    tpl = "{client} - {doc_type}[ to {recipient}]"

    check("optional segment renders when the field is present",
          NT.build(tpl, {"client": "VALADEZ, Secilia", "doc_type": "Reduction Request",
                         "recipient": "Chiropractic Works"}),
          "VALADEZ, Secilia - Reduction Request to Chiropractic Works")

    # The whole "[ to {recipient}]" group drops, not just the value — no
    # dangling " to" left behind.
    check("optional segment drops whole when the field is empty",
          NT.build(tpl, {"client": "VALADEZ, Secilia", "doc_type": "Reduction Request",
                         "recipient": ""}),
          "VALADEZ, Secilia - Reduction Request")

    check("ISO date reformatted", NT.format_date("2026-07-15", "%m-%d-%y"), "07-15-26")
    # An unparseable date is passed through, never dropped.
    check("non-ISO date passed through", NT.format_date("July 15 2026", "%m-%d-%y"),
          "July 15 2026")
    check("empty date stays empty", NT.format_date("", "%m-%d-%y"), "")

    check("illegal filename characters removed",
          NT.build("{client} - {doc_type}", {"client": "SMITH, John",
                                             "doc_type": "Report: A|B"}),
          "SMITH, John - Report AB")

    # Windows has a hard path limit; an over-long doc_type must be trimmed
    # rather than produce a name that fails to write.
    long_name = NT.build("{client} - {doc_type}",
                         {"client": "SMITH, John", "doc_type": "X" * 300})
    check("over-long name truncated to the cap", len(long_name) <= NT.MAX_LEN, True)

    # A client name keeps its exact "LAST, First" capitalisation and comma —
    # build() must not title-case it the way _safe_subject does.
    check("client capitalisation preserved",
          NT.build("{client} - {doc_type}", {"client": "O'BRIEN, Mary-Jane",
                                             "doc_type": "Lien"}),
          "O'BRIEN, Mary-Jane - Lien")


# ─────────────────────────────────────────────────────────────
# Filename construction helpers
# ─────────────────────────────────────────────────────────────

def test_safe_subject():
    section("FileProcessor._safe_subject")
    FP = sd.FileProcessor
    check("known acronym not split or lowercased", FP._safe_subject("PPR report"),
          "PPR Report")
    check("slash becomes a space rather than fusing words",
          FP._safe_subject("EMG/NCV"), "EMG NCV")
    check("underscore becomes a space",
          FP._safe_subject("Incoming_Document"), "Incoming Document")
    check("run-on words split", FP._safe_subject("RetainerAgreement"),
          "Retainer Agreement")
    check("joining words stay lowercase mid-phrase",
          FP._safe_subject("Reduction Request To Chiropractic Works"),
          "Reduction Request to Chiropractic Works")


def test_already_processed():
    section("FileProcessor._already_processed")
    FP = sd.FileProcessor
    clients = ["SMITH, John", "DOE, Jane"]
    check("a properly named file is recognised",
          FP._already_processed("SMITH, John - PPR.pdf", clients, {}), True)
    check("a raw scan is not", FP._already_processed("scan001.pdf", clients, {}), False)
    # Both unresolved labels must stay eligible for reprocessing forever,
    # otherwise a file the tool couldn't name would be skipped for good.
    check("A-NEEDS REVIEW stays eligible",
          FP._already_processed("A-NEEDS REVIEW - Unknown.pdf", clients, {}), False)
    check("A-UNKNOWN CLIENT stays eligible",
          FP._already_processed("A-UNKNOWN CLIENT - Unknown.pdf", clients, {}), False)
    # A firm can rename those labels in Settings; honour the custom text too.
    check("a customised label stays eligible",
          FP._already_processed("ZZ-UNFILED - Unknown.pdf", clients,
                                {"no_client_label": "ZZ-UNFILED"}), False)


def test_resolve_collision():
    section("FileProcessor._resolve_collision")
    FP = sd.FileProcessor
    tmp = tempfile.mkdtemp()
    try:
        # In preview mode nothing lands on disk, so two documents that would
        # take the same name must still be shown as distinct. That is what
        # `reserved` is for.
        reserved = set()
        first = FP._resolve_collision(tmp, "SMITH, John - PPR.pdf", "a.pdf",
                                      reserved=reserved)
        reserved.add(first)
        second = FP._resolve_collision(tmp, "SMITH, John - PPR.pdf", "b.pdf",
                                       reserved=reserved)
        check("two previewed files get distinct names", first != second, True)

        # naming.date_disambiguation prefers a meaningful suffix over "(1)".
        reserved = {"SMITH, John - PPR.pdf"}
        dated = FP._resolve_collision(
            tmp, "SMITH, John - PPR.pdf", "c.pdf", reserved=reserved,
            extra={"date_disambiguation": True, "doc_date": "2026-07-15",
                   "claim_number": "", "date_format": "%m-%d-%y"},
        )
        check("date used as the disambiguator", dated, "SMITH, John - PPR 07-15-26.pdf")

        # Renaming a file to the name it already has is never a collision.
        check("source name is not a collision with itself",
              FP._resolve_collision(tmp, "x.pdf", "x.pdf"), "x.pdf")
    finally:
        shutil.rmtree(tmp)


# ─────────────────────────────────────────────────────────────
# LearningStore
# ─────────────────────────────────────────────────────────────

def test_learning_store_promotion():
    section("LearningStore — alias promotion")
    tmp = tempfile.mkdtemp()
    try:
        store = sd.LearningStore(tmp, observations_required=3)
        for i in range(3):
            store.log_correction({
                "doc_hash": f"hash{i}", "original_name": f"f{i}.pdf",
                "predicted_client": "A-NEEDS REVIEW", "raw_client": "George Martinez",
                "corrected_client": "SMITH, Mary", "corrected_desc": "Medical Bills",
                "changed_client": True, "changed_desc": True,
                "claim_number": "WC-123", "source": "correction",
            })

        suggestions = store.pending_suggestions()
        check("alias surfaces once enough distinct documents agree",
              any(s["kind"] == "alias" and s["resolved_client"] == "SMITH, Mary"
                  for s in suggestions), True)
        # Promotion produces a SUGGESTION only. Nothing is applied to a real
        # file until a human confirms it.
        check("nothing is applied before a human confirms",
              store.lookup_alias("George Martinez"), "")
        store.confirm_alias("George Martinez", "SMITH, Mary")
        check("applied after confirmation",
              store.lookup_alias("George Martinez"), "SMITH, Mary")
        check("claim number resolves to the same client",
              store.lookup_claim("WC-123"), "SMITH, Mary")
    finally:
        shutil.rmtree(tmp)


def test_learning_store_guards():
    section("LearningStore — evidence guards")
    tmp = tempfile.mkdtemp()
    try:
        store = sd.LearningStore(tmp, observations_required=3)

        # Guard: a bare first name can never become an alias. Common first
        # names recur constantly in this client base.
        for i in range(4):
            store.log_correction({
                "doc_hash": f"a{i}", "predicted_client": "", "raw_client": "George",
                "corrected_client": "SMITH, Mary", "changed_client": True,
                "source": "correction",
            })
        check("a bare first name is never promoted",
              any(s["kind"] == "alias" for s in store.pending_suggestions()), False)

        # Guard: the same document logged repeatedly is one piece of
        # evidence, not many. (This has happened in production.)
        for _ in range(5):
            store.log_correction({
                "doc_hash": "same-hash", "predicted_client": "",
                "raw_client": "Robert Vance", "corrected_client": "SMITH, Mary",
                "changed_client": True, "source": "correction",
            })
        check("the same document logged 5x is not 5 observations",
              any(s.get("raw_name") == "Robert Vance"
                  for s in store.pending_suggestions()), False)

        # Guard: disagreement means ambiguous, never a suggestion. George may
        # be a passenger in two different accidents, or be two people.
        for i, client in enumerate(["SMITH, Mary", "JONES, Bill", "LEE, Ann"]):
            store.log_correction({
                "doc_hash": f"amb{i}", "predicted_client": "",
                "raw_client": "Chris Taylor", "corrected_client": client,
                "changed_client": True, "source": "correction",
            })
        check("documents that disagree produce no suggestion",
              any(s.get("raw_name") == "Chris Taylor"
                  for s in store.pending_suggestions()), False)
    finally:
        shutil.rmtree(tmp)


def test_learning_store_rejects_sentinels():
    section("LearningStore — placeholders are never clients")
    # Regression: browsing result rows with the Manual Correction panel open
    # committed each row unchanged, recording corrected_client
    # "A-NEEDS REVIEW". rebuild_index only screened the *predicted* side, so
    # the placeholder accumulated as alias evidence and was offered as
    # 'file George Martinez under A-NEEDS REVIEW'. Accepting that renamed
    # real documents to the placeholder while reporting them as renamed.
    tmp = tempfile.mkdtemp()
    try:
        store = sd.LearningStore(tmp, observations_required=3,
                                 sentinel_labels=["ZZ-UNFILED"])
        for i in range(4):
            store.log_correction({
                "doc_hash": f"h{i}", "predicted_client": "A-NEEDS REVIEW",
                "raw_client": "George Martinez",
                "corrected_client": "A-NEEDS REVIEW",
                "changed_client": False, "claim_number": "WC-999",
                "source": "confirmation",
            })
        check("a placeholder is never suggested as a client",
              [s for s in store.pending_suggestions() if s["kind"] == "alias"], [])
        check("a placeholder never enters the claim index",
              store.lookup_claim("WC-999"), "")

        # A firm's own customised label is a placeholder too.
        for i in range(4):
            store.log_correction({
                "doc_hash": f"c{i}", "predicted_client": "",
                "raw_client": "Dana Fields", "corrected_client": "ZZ-UNFILED",
                "changed_client": True, "source": "correction",
            })
        check("a customised placeholder label is also rejected",
              [s for s in store.pending_suggestions() if s["kind"] == "alias"], [])

        # Confirming one directly must be refused, not merely un-suggested.
        store.confirm_alias("George Martinez", "A-NEEDS REVIEW")
        check("confirming a placeholder is refused",
              store.lookup_alias("George Martinez"), "")

        # A real client alongside them still works.
        for i in range(3):
            store.log_correction({
                "doc_hash": f"r{i}", "predicted_client": "",
                "raw_client": "Nina Alvarez", "corrected_client": "SMITH, Mary",
                "changed_client": True, "source": "correction",
            })
        check("a real client is still promoted normally",
              any(s.get("raw_name") == "Nina Alvarez"
                  for s in store.pending_suggestions()), True)
    finally:
        shutil.rmtree(tmp)


def test_learning_store_scan_cap():
    section("LearningStore — few-shot scan is bounded")
    # find_similar_corrections runs a difflib ratio per entry and is called
    # once per document during a batch. Unbounded it measured ~5s per file
    # at 10k log entries, getting worse every run.
    tmp = tempfile.mkdtemp()
    try:
        store = sd.LearningStore(tmp)
        with open(store.log_path, "w", encoding="utf-8") as f:
            for i in range(store.MAX_FEWSHOT_SCAN + 500):
                f.write(json.dumps({
                    "text_excerpt": f"document number {i} about a knee injury",
                    "corrected_client": "SMITH, Mary", "corrected_desc": "Medical Bills",
                }) + "\n")
        results = store.find_similar_corrections("knee injury", limit=3)
        check("still returns matches", len(results) > 0, True)
        check("scan cap is well under the log size",
              store.MAX_FEWSHOT_SCAN < store.MAX_LOG_ENTRIES, True)
    finally:
        shutil.rmtree(tmp)


def test_learning_store_trim():
    section("LearningStore — log trimming")
    tmp = tempfile.mkdtemp()
    try:
        store = sd.LearningStore(tmp)
        store.MAX_LOG_ENTRIES = 100
        store.TRIM_TO_ENTRIES = 60
        with open(store.log_path, "w", encoding="utf-8") as f:
            for i in range(150):
                f.write(json.dumps({"doc_hash": f"h{i}", "source": "observation"}) + "\n")
        check("trim reports that it ran", store._trim_log_if_needed(), True)
        check("newest entries kept", len(store._read_all_corrections()), 60)
        check("trim is a no-op when under the cap", store._trim_log_if_needed(), False)
    finally:
        shutil.rmtree(tmp)


# ─────────────────────────────────────────────────────────────
# Vocabularies
# ─────────────────────────────────────────────────────────────

def test_document_type_normalize():
    section("DocumentTypeManager.normalize")
    tmp = tempfile.mkdtemp()
    try:
        path = os.path.join(tmp, "document_types.txt")
        aliases = sd.DocumentTypeManager.load_alias_map(path)   # seeds the file
        check("a known alias maps to its canonical name",
              sd.DocumentTypeManager.normalize("Compromise Offer", aliases),
              "Reduction Request")
        # The model often pads the real title with invented description.
        check("a title padded with extra words still matches",
              sd.DocumentTypeManager.normalize(
                  "Compromise Offer Letter Medical Bill", aliases),
              "Reduction Request")
        # Nothing close enough returns "" so the caller keeps the model's
        # own printed title rather than discarding it.
        check("an unknown type returns empty",
              sd.DocumentTypeManager.normalize("Wombat Certificate", aliases), "")
        # Short keys must not match inside unrelated titles.
        check("a short key does not match inside an unrelated title",
              sd.DocumentTypeManager.normalize("Alien Registration Notice", aliases),
              "")
    finally:
        shutil.rmtree(tmp)


def test_document_type_save_preserves_aliases():
    section("DocumentTypeManager.save — aliases survive a rewrite")
    # Regression: save() was a flat rewrite of canonical names only, so the
    # first press of the Document Types tab's Save button — or the first
    # accepted doc-type suggestion — silently destroyed every alias in the
    # file. normalize() matches against those aliases, so "Compromise Offer"
    # stopped resolving to "Reduction Request" with no error, no log line,
    # and the same names still listed in the tab.
    DTM = sd.DocumentTypeManager
    tmp = tempfile.mkdtemp()
    try:
        path = os.path.join(tmp, "document_types.txt")
        before = DTM.load_alias_map(path)              # seeds the file
        check("seed file carries more aliases than canonical names",
              len(before) > len(DTM.load(path)), True)

        # Exactly what the Save button and Accept-suggestion both do.
        names = DTM.load(path)
        names.append("New Type From Suggestion")
        DTM.save(path, names)

        after = DTM.load_alias_map(path)
        check("alias count is preserved (plus the new canonical)",
              len(after), len(before) + 1)
        check("an alias still resolves after saving",
              DTM.normalize("Compromise Offer", after), "Reduction Request")
        check("the padded-title containment pass still works",
              DTM.normalize("Compromise Offer Letter Medical Bill", after),
              "Reduction Request")
        check("the new type is present", "New Type From Suggestion" in DTM.load(path), True)
        check("the comment header survives",
              open(path, encoding="utf-8").read().startswith("#"), True)

        # Removing a type takes its aliases with it — an alias cannot
        # outlive the canonical it points at — and leaves the rest alone.
        names = [n for n in DTM.load(path) if n != "Reduction Request"]
        DTM.save(path, names)
        after_removal = DTM.load_alias_map(path)
        check("a removed type is gone", "Reduction Request" in DTM.load(path), False)
        check("its aliases went with it",
              DTM.normalize("Compromise Offer", after_removal), "")
        check("an unrelated type keeps its aliases",
              DTM.normalize("Physician Progress Report", after_removal), "PPR")
    finally:
        shutil.rmtree(tmp)


def test_provider_save_preserves_header():
    section("ProviderManager.save — header survives a rewrite")
    PM = sd.ProviderManager
    tmp = tempfile.mkdtemp()
    try:
        path = os.path.join(tmp, "providers.txt")
        PM.load(path)                                   # seeds the header
        PM.save(path, ["Chiropractic Works", "First Care"])
        check("the comment header survives",
              open(path, encoding="utf-8").read().startswith("#"), True)
        check("entries round-trip",
              PM.load(path), ["Chiropractic Works", "First Care"])
        # An indented comment is a comment, not a provider name.
        with open(path, "a", encoding="utf-8") as f:
            f.write("   # indented note\n")
        check("an indented comment is not read as a provider",
              PM.load(path), ["Chiropractic Works", "First Care"])
    finally:
        shutil.rmtree(tmp)


# ─────────────────────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────────────────────

def test_config_merge():
    section("ConfigManager._deep_merge")
    CM = sd.ConfigManager
    defaults = CM._deep_copy(sd.DEFAULT_CONFIG)

    merged = CM._deep_merge(CM._deep_copy(defaults), {"api": {"model": "custom-model"}})
    check("a saved scalar overrides the default",
          merged["api"]["model"], "custom-model")
    check("keys absent from the saved file keep their defaults",
          merged["api"]["timeout_read"], defaults["api"]["timeout_read"])
    # A section from a newer version of the app must survive a downgrade.
    merged = CM._deep_merge(CM._deep_copy(defaults), {"future_section": {"x": 1}})
    check("an unknown saved section is preserved",
          merged["future_section"], {"x": 1})

    # naming.templates is user data, not a set of named settings: recursive
    # merging kept resurrecting stock rows the user had deleted, with no way
    # to be rid of them.
    merged = CM._deep_merge(CM._deep_copy(defaults),
                            {"naming": {"templates": {"PPR": "{client} - {doc_type}"}}})
    check("a deleted template stays deleted",
          merged["naming"]["templates"], {"PPR": "{client} - {doc_type}"})
    check("sibling naming settings still merge normally",
          merged["naming"]["date_format"], defaults["naming"]["date_format"])


def test_settings_dependencies():
    section("compute_settings_dependency_state")
    cfg = sd.ConfigManager._deep_copy(sd.DEFAULT_CONFIG)
    state = sd.compute_settings_dependency_state(cfg)
    disabled, reason = state["classification.use_document_types"]
    check("dependent control disabled while its prerequisite is off", disabled, True)
    check("and the reason is explained", bool(reason), True)

    cfg["classification"]["structured_output"] = True
    state = sd.compute_settings_dependency_state(cfg)
    check("enabled once the prerequisite is on",
          state["classification.use_document_types"], (False, ""))

    # Vision escalation needs a model that can actually see images.
    cfg["api"]["model"] = "llama3.2"
    check("vision escalation gated on a non-vision model",
          sd.compute_settings_dependency_state(cfg)["reading.vision_escalation"][0], True)
    cfg["api"]["model"] = "llama3.2-vision"
    check("vision escalation available on a vision model",
          sd.compute_settings_dependency_state(cfg)["reading.vision_escalation"][0], False)
    # Cloud variants send documents off-device and are excluded on purpose.
    cfg["api"]["model"] = "llama3.2-vision-cloud"
    check("cloud vision variants are excluded",
          sd.compute_settings_dependency_state(cfg)["reading.vision_escalation"][0], True)


def test_mode_round_trip():
    section("Settings display <-> config mapping")
    for value in ("off", "suggest", "auto"):
        check(f"learning mode {value!r} round-trips",
              sd.learning_mode_to_config(sd.learning_mode_to_display(value)), value)
    # A hand-edited config must never crash the UI.
    check("an unrecognised learning mode falls back to Off",
          sd.learning_mode_to_display("nonsense"), "Off")
    for value in ("off", "preview"):
        check(f"retroactive mode {value!r} round-trips",
              sd.retroactive_mode_to_config(sd.retroactive_mode_to_display(value)), value)
    # There is deliberately no "Automatic" here: a law firm must never have
    # files renamed in bulk with no human in the loop.
    check("retroactive rename has no automatic mode",
          "Automatic" in sd.RETROACTIVE_MODE_VALUES, False)


# ─────────────────────────────────────────────────────────────
# Batch orchestration
# ─────────────────────────────────────────────────────────────

def _run_fake_batch(config, folder, renamer):
    """Run ProcessingEngine.run_batch with process_file stubbed out, so the
    orchestration can be tested without OCR or a model. Returns
    (files handed to process_file, message types queued)."""
    seen = []
    original = sd.FileProcessor.process_file

    def fake(file_path, cfg, client_list, batch_id="", reserved=None):
        name = os.path.basename(file_path)
        seen.append(name)
        new_name, status, client = renamer(name)
        if new_name != name and os.path.isfile(file_path):
            os.rename(file_path, os.path.join(os.path.dirname(file_path), new_name))
        return sd.ProcessResult(original_name=name, final_name=new_name,
                                status=status, client=client)

    sd.FileProcessor.process_file = staticmethod(fake)
    try:
        q = queue.Queue()
        sd.ProcessingEngine().run_batch(config, q)
        types_seen = []
        while not q.empty():
            types_seen.append(q.get()["type"])
        return seen, types_seen
    finally:
        sd.FileProcessor.process_file = original


def _batch_config(folder, clients=("SMITH, John",)):
    cfg = sd.ConfigManager._deep_copy(sd.DEFAULT_CONFIG)
    cfg["paths"]["scandocs_folder"] = folder
    cfg["paths"]["client_list_file"] = os.path.join(folder, "clients.txt")
    with open(cfg["paths"]["client_list_file"], "w", encoding="utf-8") as f:
        f.write("\n".join(clients))
    cfg["safety"]["instance_lock"] = False
    return cfg


def test_straggler_sweep_no_reprocessing():
    section("run_batch — straggler sweep")
    # Regression: the sweep compared a fresh listing against the ORIGINAL
    # filenames, so every file the batch had just renamed looked newly
    # arrived. _already_processed was supposed to filter those out, but it
    # deliberately returns False for the unresolved labels — so every
    # needs-review file was read, classified and billed for a second time,
    # and appeared twice in the report. With skip_already_processed off,
    # the entire batch ran twice.
    def renamer(name):
        if name.startswith("scan001"):
            return "SMITH, John - Reduction Request.pdf", "renamed", "SMITH, John"
        return f"A-NEEDS REVIEW - Unknown {name[4:7]}.pdf", "needs_review", "A-NEEDS REVIEW"

    for skip_flag in (True, False):
        tmp = tempfile.mkdtemp()
        try:
            for n in ("scan001.pdf", "scan002.pdf", "scan003.pdf"):
                with open(os.path.join(tmp, n), "w") as f:
                    f.write("x")
            cfg = _batch_config(tmp)
            cfg["processing"]["skip_already_processed"] = skip_flag
            seen, _ = _run_fake_batch(cfg, tmp, renamer)
            check(f"each file processed exactly once (skip_already_processed={skip_flag})",
                  sorted(seen), ["scan001.pdf", "scan002.pdf", "scan003.pdf"])
        finally:
            shutil.rmtree(tmp)


def test_straggler_sweep_still_picks_up_new_files():
    section("run_batch — genuinely new files are still collected")
    tmp = tempfile.mkdtemp()
    try:
        with open(os.path.join(tmp, "scan001.pdf"), "w") as f:
            f.write("x")
        cfg = _batch_config(tmp)

        dropped = {"done": False}

        def renamer(name):
            # Simulate a second workstation dropping a file in mid-batch.
            if not dropped["done"]:
                dropped["done"] = True
                with open(os.path.join(tmp, "scan999.pdf"), "w") as f:
                    f.write("y")
            return f"SMITH, John - Doc {name[4:7]}.pdf", "renamed", "SMITH, John"

        seen, _ = _run_fake_batch(cfg, tmp, renamer)
        check("a file that arrived mid-batch is picked up",
              sorted(seen), ["scan001.pdf", "scan999.pdf"])
    finally:
        shutil.rmtree(tmp)


def test_batch_queues_exactly_one_done():
    section("run_batch — one 'done' per run")
    # Regression: early-error paths queued "done" explicitly AND again in
    # the finally. _poll_queue stops reading at "error", so the extra stayed
    # in the queue and the NEXT run consumed it immediately — reporting
    # "Done. 0 file(s) processed" and stopping the poll loop while the batch
    # it had just started kept renaming files in the background.
    tmp = tempfile.mkdtemp()
    try:
        cfg = _batch_config(tmp, clients=())          # empty client list
        seen, types_seen = _run_fake_batch(cfg, tmp, lambda n: (n, "renamed", ""))
        check("empty client list reports an error", "error" in types_seen, True)
        check("exactly one 'done' is queued", types_seen.count("done"), 1)
    finally:
        shutil.rmtree(tmp)

    tmp = tempfile.mkdtemp()
    try:
        cfg = _batch_config(tmp)                      # no documents at all
        _, types_seen = _run_fake_batch(cfg, tmp, lambda n: (n, "renamed", ""))
        check("empty folder queues exactly one 'done'", types_seen.count("done"), 1)
    finally:
        shutil.rmtree(tmp)


def test_batch_config_is_snapshotted():
    section("run_batch — config snapshot")
    # Nothing stops the user hitting Save in Settings while a batch runs.
    # Without a snapshot that changes the rules mid-batch — flip
    # automation.dry_run and half the run previews while half really renames.
    tmp = tempfile.mkdtemp()
    try:
        with open(os.path.join(tmp, "scan001.pdf"), "w") as f:
            f.write("x")
        cfg = _batch_config(tmp)
        seen_dry_run = []

        original = sd.FileProcessor.process_file

        def fake(file_path, config, client_list, batch_id="", reserved=None):
            seen_dry_run.append(config["automation"]["dry_run"])
            # Simulate the user saving Settings mid-batch.
            cfg["automation"]["dry_run"] = True
            return sd.ProcessResult(original_name=os.path.basename(file_path),
                                    final_name=os.path.basename(file_path),
                                    status="renamed")

        sd.FileProcessor.process_file = staticmethod(fake)
        try:
            sd.ProcessingEngine().run_batch(cfg, queue.Queue())
        finally:
            sd.FileProcessor.process_file = original

        check("the batch does not see a mid-run config change",
              seen_dry_run, [False])
    finally:
        shutil.rmtree(tmp)


# ─────────────────────────────────────────────────────────────
# Undo log
# ─────────────────────────────────────────────────────────────

def test_rename_log_undo():
    section("RenameLog — undo")
    tmp = tempfile.mkdtemp()
    try:
        log_path = os.path.join(tmp, "renames.jsonl")
        src = os.path.join(tmp, "original.pdf")
        dst = os.path.join(tmp, "SMITH, John - PPR.pdf")
        with open(src, "w") as f:
            f.write("x")
        os.rename(src, dst)
        sd.log_rename("batch1", "rename", src, dst, "auto", log_path=log_path)

        log = sd.RenameLog(log_path)
        check("the batch is listed", [b["batch_id"] for b in log.last_batches()], ["batch1"])
        undone, skipped, errors = log.undo_batch("batch1")
        check("one rename undone", undone, 1)
        check("no errors", errors, [])
        check("the file is back at its original name", os.path.isfile(src), True)

        # Undoing twice must not error — the file is simply no longer where
        # the log says it was left.
        undone, skipped, _ = log.undo_batch("batch1")
        check("a second undo is a safe no-op", (undone, skipped), (0, 1))
    finally:
        shutil.rmtree(tmp)


# ─────────────────────────────────────────────────────────────

def main():
    if STUBBED:
        print("note: tkinter/ttkbootstrap not installed — using import stubs "
              "(UI code is not exercised by these tests either way)")

    for test in (
        test_name_template,
        test_safe_subject,
        test_already_processed,
        test_resolve_collision,
        test_learning_store_promotion,
        test_learning_store_guards,
        test_learning_store_rejects_sentinels,
        test_learning_store_scan_cap,
        test_learning_store_trim,
        test_document_type_normalize,
        test_document_type_save_preserves_aliases,
        test_provider_save_preserves_header,
        test_config_merge,
        test_settings_dependencies,
        test_mode_round_trip,
        test_straggler_sweep_no_reprocessing,
        test_straggler_sweep_still_picks_up_new_files,
        test_batch_queues_exactly_one_done,
        test_batch_config_is_snapshotted,
        test_rename_log_undo,
    ):
        test()
        print(f"   {test.__name__}")

    total = _PASSED + len(_FAILED)
    print("\n" + "=" * 62)
    if _FAILED:
        print(f"FAILED — {len(_FAILED)} of {total} checks\n")
        for failure in _FAILED:
            print(f"  ✗ {failure}")
        return 1
    print(f"PASSED — all {total} checks")
    return 0


if __name__ == "__main__":
    sys.exit(main())
