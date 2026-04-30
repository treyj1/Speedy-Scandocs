Subject: Speedy Scandocs - Mac Installation Instructions

Hi Brig,

Here's how to get Speedy Scandocs installed on the office Macs.

---

## Install Speedy Scandocs

1. Download **SpeedyScandocs.dmg** from:
   https://github.com/treyj1/Speedy-Scandocs/releases/tag/v1.7

2. Double-click the DMG to open it
3. Drag **SpeedyScandocs** into your **Applications** folder
4. Open it from Applications

**First Launch - macOS will block it** because it's not from the App Store:
   - Go to **System Settings > Privacy & Security**
   - Scroll down - you'll see a message about SpeedyScandocs being blocked
   - Click **Open Anyway**
   - You only need to do this once

If it still won't open after that, reach out to me and I can help.

## Configure the App

Open the **Settings** tab and set:

| Setting | Value |
|---------|-------|
| **Scandocs Folder** | Browse to the folder where scanned documents land |
| **Client List File** | Browse to the client_list.txt file |
| **Model** | `ministral-3:8b` |
| **Fuzzy Match Threshold** | `0.85` |
| **Max OCR Characters** | `8000` |
| **Max Pages Per Document** | `5` |
| **Require High Confidence** | Checked |
| **Audit Mode** | Checked |
| **Report Folder** | Browse to wherever you want reports saved |

Set the **OpenWebUI URL** and **Ollama URL** to the address where you access OpenWebUI and Ollama on your network. Click **Save** at the bottom.

---

## How to Use

1. Make sure **Ollama is running**
2. Open **Speedy Scandocs**
3. Click **Auto-Process Documents**
4. Review results - press **Spacebar** to preview a document, arrow keys to navigate
5. Mark each result as Correct or flag any issues
6. Click **Submit Audit** - the report and any flagged files save to your report folder

Let me know if you run into any issues!
