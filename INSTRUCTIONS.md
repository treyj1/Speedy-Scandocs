# Speedy Scandocs — How to Use

A short, plain-English guide to running Speedy Scandocs day-to-day.

## Before your first run

### 1. Get the client list from Trey

The client list is **not bundled with the app**. Trey ([treyjensen@icloud.com](mailto:treyjensen@icloud.com)) will send the current list to you.

When you receive it:

1. Save the file as `client_list.txt` (plain text, one client per line in `LAST, First` format).
2. Put it **inside the same Scandocs folder you use every day** for scanning — the folder where the scanner drops PDFs/JPEGs. Keeping the list and the documents in one place means fewer paths to remember and fewer ways to get out of sync.
3. Open Settings in the app and point both **Scandocs Folder** and **Client List File** at that folder (the client list field should be the full path to `client_list.txt` inside it).

Whenever Trey sends an updated list, replace the file in the same spot. Don't rename it and don't move it — just overwrite.

### 2. Pick a model

In Settings → Model, the recommended choice is:

> **`ministral-3:8b`** — the best model tested so far. Good balance of speed and accuracy on real scanned documents.

(You may also see it referred to as `ministral-3:7b` in older notes — the 8b variant is the one to use.)

Newer models may eventually beat it. If you want to try something else, click the refresh button next to the Model field to see what's installed on the Ollama server, swap it in, and run a small batch through audit mode to compare. If accuracy drops, switch back to `ministral-3:8b`.

Cloud models (anything not running locally on Ollama) are intentionally excluded — client documents must stay on the local machine.

## Daily workflow

1. **Scan documents** into your Scandocs folder as usual.
2. **Open Speedy Scandocs.**
3. Click **Auto-Process Documents.** The app will read each PDF/JPEG, identify the client, and rename the file to `CLIENT NAME - Description.pdf`.
4. **Review the audit panel.** For each result, mark it:
   - **Correct** — rename was right
   - **Wrong Client Name** — file will be flagged "A-NEEDS REVIEW" on submit
   - **Bad Description** — description gets reset to "Scanned Document"
   - **Failed to Identify Client** — tool couldn't find a match
   - **Should Have Been Flagged** — borderline case the tool shouldn't have auto-renamed
5. Click **Submit Audit.** This applies corrections, saves an Excel report to the Reports folder, and copies any wrong-client files there for follow-up.
6. (Optional) Use **File Mode** to move the renamed files to each client's individual folder.

Press **Spacebar** on a selected file to preview it. Arrow keys move between files.

## Known pitfall: employee names on the client list

If anyone on your client list is **also an employee** (or otherwise appears as a name *inside* documents that aren't theirs — e.g. a paralegal whose name is in the letterhead, or a staff member CC'd on correspondence), the AI may see their name on the page and incorrectly tag them as the client.

What to do about it:

- **Don't put employees on the client list** unless they are also an actual client. The list is for clients, not staff.
- If an employee genuinely is a client, expect more wrong-client flags on their documents and watch for it during audit.
- During audit, double-check any document where the matched name is a known employee — that's the most likely place a false match will show up.
- If you see a recurring false match (the AI keeps tagging Employee X as the client across unrelated documents), tell Trey so the client list or matching threshold can be adjusted.

## Settings cheat-sheet

These are the recommended values. Don't change them unless you know what you're doing.

| Setting | Value |
|---|---|
| Model | `ministral-3:8b` |
| Fuzzy Match Threshold | `0.85` |
| Max OCR Characters | `8000` |
| Max Pages Per Document | `5` |
| Require High Confidence | On |
| Skip Already Processed | On |
| Audit Mode | On |

## When something goes wrong

- **App can't find clients / every file goes to "needs review"** — Settings → Client List File path is probably wrong, or the file is empty. Check that `client_list.txt` is in the Scandocs folder and the path in Settings matches.
- **App says it can't reach the AI** — Ollama isn't running. Start Ollama and try again.
- **Results are noticeably worse than usual** — confirm the model is still set to `ministral-3:8b` (it can get changed accidentally in Settings).
- **Anything else** — email Trey ([treyjensen@icloud.com](mailto:treyjensen@icloud.com)) with the file name and a screenshot.
