# Known gaps / follow-up work

Deliberate deferrals, not bugs. Each entry says what's missing, why it was
left, and what a fix would involve, so picking one up later doesn't mean
re-deriving the context.

---

## No UI for editing filename templates

**Status:** deferred — raised during the v2.0.0 review, held for a product decision.

The per-document-type filename template engine (`NameTemplate`, and the
`naming.templates` / `naming.default_template` config) is implemented,
tested and working. Its on/off switch — "Use document-type naming
templates" — is in Settings. But the templates themselves have no editor:
they can only be changed by hand-editing `config.json`.

So the headline naming feature of the 2.0 work is half-reachable. A user
can turn it on and get the six stock templates, but can't add one for a
document type the firm actually sees, or change how an existing one reads.

**The workaround that exists today:** Settings → Export Settings, edit the
JSON, Settings → Import Settings. Works, and `naming.templates` is now
replaced wholesale on load, so a template deleted that way stays deleted
instead of reappearing on the next start.

**What a fix has to decide:**

- Where it lives. Its own tab, or a sub-panel under the existing Naming
  section in Settings?
- What the editing surface is. A per-doc-type row bound to the Document
  Types list, or a free-form key/value table?
- How a template is validated before saving. `{placeholder}` names are a
  fixed set (`client`, `doc_type`, `recipient`, `doc_date`,
  `claim_number`, `direction`) and `[optional segments]` don't nest — both
  are worth checking at entry rather than at rename time.
- Whether to show a live preview. Rendering the template against a sample
  record as the user types is probably the single thing that makes the
  bracket syntax explain itself.
- What happens to a document type with no template — today it silently
  falls back to `naming.default_template`, which is right, but the UI
  should say so rather than leave a blank row looking broken.

**Relevant code:** `NameTemplate` (render/build and the template language
docstring above it), the filename-construction block in
`FileProcessor.process_file` that selects a template per doc type, and
`_build_settings_tab`'s `register()` pattern — note that the registry
round-trips plain scalars only, so a templates editor needs explicit
load/save handling like `s_extraction_method_var` has.

**Related:** `naming.date_format`, `naming.unknown_client_label` and
`naming.no_client_label` are in the same position — used by the naming
pipeline, no UI. They're plain strings, so they'd be cheap to add to the
existing registry if a templates editor lands.
