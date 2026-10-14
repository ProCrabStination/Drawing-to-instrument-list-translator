# Drawing to Instrument List Translator

Turns drawing packages for controls and mechanical (P&IDs) into an instrument list and a filled-in cable calculations workbook.

A vision model reads each drawing page (together with OCR text of the page), lists the PLC-connected instruments, and the tool then merges the controls and P&ID results into one combined list. The last step writes that list into a copy of your `Instrument Cable Calculations.xlsx` workbook.

## Contents

- [How it works](#how-it-works)
- [Requirements](#requirements)
- [Installation](#installation)
- [Choosing a model backend](#choosing-a-model-backend)
- [Using the main tool (GUI)](#using-the-main-tool-gui)
- [Using the main tool (command line)](#using-the-main-tool-command-line)
- [Output files](#output-files)
- [The Cable Calculations workbook step](#the-cable-calculations-workbook-step)
- [Schematic to Table Translator (companion tool)](#schematic-to-table-translator-companion-tool)
- [Other scripts](#other-scripts)
- [Configuration and environment variables](#configuration-and-environment-variables)
- [Troubleshooting](#troubleshooting)
- [Project layout](#project-layout)

## How it works

A full run has three stages:

1. **Controls stage** - reads the controls (electrical/PLC) drawing PDF page by page and lists each PLC instrument with its description, responsibility, wire count, wire type (shielded or unshielded), intrinsic-safety flag and notes.
2. **P&ID stage** - does the same for the P&ID PDF.
3. **Combine stage** - pairs up instruments that appear in both sets, merges their notes, and produces one combined list.

Optionally, a **workbook step** then fills the combined list into a copy of the Cable Calculations workbook (no model involved).

Before each page goes to the model, the PDF is cleaned (deskewed and upsampled) and OCR'd with Docling. The OCR text is sent along with the page image, which makes small tag numbers much more reliable. OCR results are cached in `ocr_cache/`, so re-runs are fast.

Two reference images (`AI Reference - Controls - smaller.png` and `AI Reference - P&IDs - smaller.png`) are sent with each page as a notation legend, so the model knows the symbols used in your drawings.

## Requirements

| Requirement | Needed for | Notes |
|---|---|---|
| Windows 10/11 | Everything | The workbook step and the launcher are Windows-specific |
| Python 3.11+ (with Tkinter) | Everything | Tkinter ships with the standard python.org installer |
| A model backend | Everything | One of: Claude Code CLI, Ollama, or a cloud API key (see [Choosing a model backend](#choosing-a-model-backend)) |
| Tesseract OCR | PDF cleaning step | Must be on `PATH`. Without it cleaning is skipped and the original PDF is used |
| Ghostscript | PDF cleaning step | Required by `ocrmypdf` |
| Microsoft Excel | Workbook step | Driven through PowerShell/COM so the template's images, tables and validation are preserved |
| LibreOffice | Only for `Docling_Convert_Ollama.py` | Expected at `C:\Program Files\LibreOffice\program\soffice.exe` |
| `pngquant`, `unpaper` | Optional | `pngquant` shrinks cleaned PDFs; `unpaper` is only used for deep cleaning |

## Installation

**Install the Python packages** into a `.venv` folder in the project (the launcher expects that name and location):

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

   The first OCR run downloads Docling/RapidOCR model files, so allow a few minutes and an internet connection the first time.

**Launch the tool** - double-click `Launch.bat`.

Before your first run, set up a model backend (see the next section).

## Choosing a model backend

The model picker in the GUI lists everything that is available. Pick one of these:

### Claude Code CLI (default choice)

Models appear as **`Claude Code: <model>`** (`default`, `sonnet`, `opus`, `haiku`, plus explicit model IDs). Each page is one headless `claude -p` call, using your existing Claude login - no API key.

1. Install Claude Code and make sure `claude` is on `PATH` (`claude --version`).
2. Log in once by running `claude` in a terminal.
3. Restart the GUI (or press **Refresh list**). The `Claude Code: ...` models now appear.

`Claude Code: default` lets the CLI pick its own default model.

### Ollama (local, free, offline)

1. Install [Ollama](https://ollama.com) and make sure it is running (it listens on `http://localhost:11434`).
2. Pull a vision-capable model, for example:

   ```powershell
   ollama pull qwen3-vl:4b-instruct
   ollama pull gemma3:4b
   ```

3. Press **Refresh list** in the GUI. Local models are discovered automatically and need no API key.

### Cloud APIs

Anthropic, OpenAI and Google models can be used with an API key. Enter the key in the **API key** box (tick **Show** to check it) and press **Refresh list**. The host is chosen automatically for each provider.

## Using the main tool (GUI)

Launch with `Launch.bat` (no console window) or by running `Instrument_List_Transcriber.py` with no arguments. Settings are saved automatically after each change and restored next time (`instrument_list_settings.json`).

### 1. Mode

| Mode | What it does |
|---|---|
| **Full run** | Controls + P&ID stages, then combines them (and fills the workbook if enabled) |
| **Controls drawings only** | Runs just the controls stage |
| **P&ID drawings only** | Runs just the P&ID stage |
| **Combine two saved results** | Merges a controls `.json` and a P&ID `.json` from earlier runs |

### 2. Drawings

- **Controls PDF** / **P&ID PDF** - press **Browse...** or paste a path. The last few PDFs you used are remembered in the drop-down.
- **Pages** - leave empty for every page, or enter a selection such as `1,3-5`. Useful for testing on a few pages first.
- In **Combine** mode this section becomes **Saved results to combine**: choose the controls and P&ID `.json` files instead.

### 3. Model

Choose the model and, for cloud APIs, enter the API key. See [Choosing a model backend](#choosing-a-model-backend). A hint under the box tells you whether the selected model is ready to run.

### 4. Output

- **Save to** - output folder. Empty means next to each PDF. File names are generated from the PDF names.
- **Also save Excel (.xlsx) copies** - writes an `.xlsx` beside each results `.json`.
- **Read the drawing text with OCR first** - recommended. Slower the first time, cached afterwards.
- **At the end, fill the Cable Calculations workbook** - runs the workbook step after combining. Your template is copied, never modified.

### 5. Run

Press **Run**. The right-hand panel shows progress, the model's live output, and the log. Afterwards:

- **Open output folder** opens where the results were written.
- **Open run log** opens this run's log (every run is also written live to `logs/`).
- **Save log...** / **Clear log** manage the on-screen log.

Closing the window mid-run stops any running model request and Excel instance cleanly.

### Advanced options

Click the **Advanced** section to expand it. Leave fields empty for automatic behaviour.

- **Different model per stage** - override the model for controls, P&IDs or the combine stage. For example, use a strong model for the drawings and a cheaper one to combine.
- **Tuning**
  - *Max tokens* - output limit per request (automatic if empty; the combine stage can need a large value for long lists). Turn this up to at least 65000 for local models.
  - *DPI* - page render resolution (automatic: the PDF's native resolution, capped at 300).
  - *OCR scale* - Docling image scale for OCR (default `3.0`; higher is sharper and slower).
  - *OCR: analyse table structure* - on by default.
  - *OCR: describe pictures (slow)* - off by default.
- **Reference images** - the built-in legends are used by default. Untick the box or choose your own images to override them per stage.
- **Cable Calculations workbook** - choose a different template workbook.

## Using the main tool (command line)

Run the script with any arguments and it behaves as a CLI instead of opening the GUI.

```powershell
# Full run: both PDFs, then combine
python Instrument_List_Transcriber.py `
    --run-mode all `
    --controls-pdf "controls.pdf" --pids-pdf "pids.pdf" `
    --model "claude-sonnet-5-5" --api-key sk-ant-... `
    --max-tokens 32000 `
    --controls-output controls.json --pids-output pids.json --combine-output combined.json

# Controls only, first five pages, local Ollama model
python Instrument_List_Transcriber.py --run-mode controls `
    --controls-pdf "controls.pdf" --controls-pages 1-5 `
    --model qwen3-vl:4b-instruct --api-host http://localhost:11434 `
    --controls-output controls.json

# Combine two saved results
python Instrument_List_Transcriber.py --run-mode combine `
    --controls-input-json controls.json --pids-input-json pids.json `
    --model claude-sonnet-5-5 --api-key sk-ant-... --combine-output combined.json
```

### Options

| Option | Purpose |
|---|---|
| `pdf_path` (positional) | Fallback PDF for any stage that has no `--controls-pdf` / `--pids-pdf` |
| `--run-mode {all,controls,pids,combine}` | Which stage(s) to run (default `all`) |
| `--model`, `--api-host`, `--api-key` | Model and endpoint used by every stage unless overridden |
| `--max-tokens` | Maximum output tokens (default 2048; use 32000 for real drawings) |
| `--dpi` | Page render resolution (default 300) |
| `--output` | Output JSON path for the combined result |
| `--reference-image` | Notation image sent with every page |
| `--controls-pdf`, `--controls-pages`, `--controls-output` | Controls stage input, page selection and output |
| `--controls-model`, `--controls-api-host`, `--controls-api-key`, `--controls-max-tokens`, `--controls-dpi`, `--controls-reference-image` | Per-stage overrides for controls |
| `--pids-pdf`, `--pids-pages`, `--pids-output` | P&ID stage input, page selection and output |
| `--pids-model`, `--pids-api-host`, `--pids-api-key`, `--pids-max-tokens`, `--pids-dpi`, `--pids-reference-image` | Per-stage overrides for P&IDs |
| `--controls-input-json`, `--pids-input-json` | Saved stage results to combine (combine mode) |
| `--combine-model`, `--combine-api-host`, `--combine-api-key`, `--combine-max-tokens`, `--combine-output` | Per-stage overrides and output for the combine stage |

Notes:

- The CLI prints each stage's JSON to the console and saves it only if an output path is given.
- The CLI does not run the workbook step; use the GUI, or call `Cable_Workbook_Step.fill_cable_workbook()` from Python (see below).
- Claude Code models can be used from the CLI by passing `--model "Claude Code: sonnet"` (and the other `Claude Code: ...` names).

## Output files

For a controls PDF named `PG-984E Rev4.pdf` and a P&ID PDF named `PG-984B PI REV 3.pdf`, a full run writes (by default, next to each PDF or in the folder chosen under **Save to**):

| File | Contents |
|---|---|
| `<name> - controls.json` / `.xlsx` | Per-page instrument list from the controls drawings |
| `<name> - pids.json` / `.xlsx` | Per-page instrument list from the P&IDs |
| `<name> - combined.json` / `.xlsx` | Merged list pairing controls and P&ID entries, with both sets of notes |
| `<name> - Cable Calculations.xlsx` | Copy of the template with the combined list filled in (if enabled) |
| `logs/<timestamp>.txt` | The full log of the run |
| `ocr_cache/` | Cached OCR text (safe to delete; it is rebuilt on demand) |

Each extracted instrument includes: `name` (tag), `description`, `Responsibility`, `wire_count`, `wire_type`, `intrinsic_safety` and `Notes`. Anything the model cannot read is marked `[unclear]` rather than guessed - review those entries by hand.

## The Cable Calculations workbook step

Done entirely in code - no model is involved. For every instrument in the combined list it builds a row:

| Column | Filled from |
|---|---|
| TAG NO | Instrument tag |
| Controls Notes / P&IDs Notes | Notes from each drawing set |
| Voltage | Voltage, if present |
| INTRINSICALLY SAFE | `YES` / `NO` from the intrinsic-safety flag |
| SIGNAL | Tag function letters: analog tags (e.g. PT, TT) give `4-20 mA`, discrete (e.g. ZSO, LSH) give `DIGITAL CONTACT` |
| CONNECTION | Wire count and type, e.g. `2 wire shielded` |
| Cable Designation / Cable Qty | Looked up in the template's own **IN STOCK CABLES** table. A shielded run uses one multi-conductor shielded cable; an unshielded run uses single-conductor wire, one per wire |

The rows are written into a **copy** of the template through Excel (PowerShell/COM), which preserves images, tables and data validation that Python spreadsheet libraries drop. Afterwards the copy is compared with the template cell by cell, and the log reports how many cells changed inside the instrument list (expected) and warns if anything changed outside it.

Requirements: Excel installed, and the template must contain the list sheet (`Instrumentation Cables`) with the expected column headings. Cable fields stay empty if no stock cable matches, or if the template has no stock-cable table (a warning is logged).

To run it on its own from Python:

```python
from pathlib import Path
import Cable_Workbook_Step as step

step.fill_cable_workbook(Path("combined.json"), Path("out.xlsx"), template=step.DEFAULT_TEMPLATE)
```

## Schematic to Table Translator (companion tool)

`Schematic_To_Table_Translator.py` is a more general companion. Instead of only the PLC instrument list, it transcribes *everything meaningful* on each page of a drawing PDF:

- components/devices (tag, type, description, location)
- connections (from, to, line style, wire label, wire count)
- any printed tables (title block, revision table, BOM, I/O list)
- free text (notes, callouts) and a plain-language summary

For accuracy, each page is first split into circuit sections (with bounding boxes); each section is cropped from the full-resolution render and translated separately, then merged. An overlay image of the boxes is saved so you can check the split by eye. Pages that cannot be split fall back to a single whole-page pass.

Run with no arguments for its GUI, or use the CLI:

```powershell
python Schematic_To_Table_Translator.py "drawing.pdf" --model claude-sonnet-5-5 `
    --api-host https://api.anthropic.com --api-key sk-ant-... --pages 1-3
```

| Option | Default | Purpose |
|---|---|---|
| `pdf` | - | Drawing PDF (required) |
| `--model` | `qwen3-vl:4b-instruct` | Model name (env `AI_GATEWAY_MODEL`) |
| `--api-host` | `http://localhost:11434` | Endpoint (env `AI_GATEWAY_HOST`) |
| `--api-key` | `None` | API key (env `AI_GATEWAY_API_KEY`) |
| `--max-tokens` | 8192 | Max output tokens per call |
| `--num-ctx` | 16384 | Ollama context window; keep modest (16384-32768) |
| `--dpi` | 300 | Page render resolution |
| `--pages` | all | Page selection, e.g. `1,3-5` |
| `--reference-image` | none | Notation legend image |
| `--no-ocr` | off | Skip the OCR pre-step |
| `--ocr-scale` | 4.0 | OCR image scale |
| `--no-sections` | off | One pass per page instead of section splitting |
| `--max-sections` | 12 | Maximum sections per page |
| `--output` | next to the PDF | Output base path |

Outputs, next to the PDF unless `--output` is given: `<name> - translated.json`, `.md` (report) and `.xlsx` (one sheet per data type), plus a `<name> - translated - sections/` folder of overlay images.

## Project layout

```
Instrument_List_Transcriber.py     Main tool: pipeline, CLI, model requests (launches the GUI with no arguments)
Instrument_List_GUI.py             The main tool's window
Pdf_OCR_Preprocess.py              PDF cleaning + Docling OCR with caching
Claude_Code_Backend.py             Claude Code CLI backend
Cable_Workbook_Step.py             Fills the Cable Calculations workbook
Docling_Convert_Ollama.py          Standalone document converter
Instrument Cable Calculations.xlsx Workbook template
AI Reference - *.png               Notation legends sent with each page
Launch.bat                         Double-click launcher
requirements.txt                   Python dependencies
```
