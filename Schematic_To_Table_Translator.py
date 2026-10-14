"""Translate the schematics in a PDF into tables and text.

Companion to Instrument_List_Transcriber.py. Where that script extracts only a
PLC instrument list, this one reads every page of a drawing PDF with a vision
model and transcribes *everything meaningful* into structured data:

  * components / devices (tag, type, description, location)
  * connections (from -> to, line style, wire label, wire count)
  * any tables printed on the drawing (title block, revision table, BOM, I/O list)
  * free text (notes, callouts) and a plain-language summary

Accuracy: each page is first split by the model into circuit sections (with bounding
boxes); every section is cropped from the full-resolution render and translated on its
own, then the results are merged. An overlay image of the boxes is saved so the split can
be checked by eye. Pages that can't be split fall back to a single whole-page pass.

Outputs: a JSON file, a Markdown report, and an Excel workbook (one sheet per
data type). It reuses the request/rendering/JSON-recovery helpers from
Instrument_List_Transcriber.py, so model selection (Ollama or Claude), API host
and key behave the same way.

Example:
    python Schematic_To_Table_Translator.py drawing.pdf --model claude-sonnet-5-5 \
        --api-host https://api.anthropic.com --api-key sk-... --pages 1-3
"""

import argparse
import datetime
import json
import os
import queue
import shutil
import sys
import tempfile
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, List, Optional

import requests
from PIL import Image, ImageDraw

from Instrument_List_Transcriber import (
    _extract_json_object,
    _get_claude_models,
    _get_local_ollama_models,
    _is_claude_model,
    _load_image_bytes,
    _normalize_host,
    openpyxl,
    parse_page_selection,
    render_pdf_pages_to_images,
    with_reference_note,
)
from Pdf_OCR_Preprocess import preprocess_pdf_ocr


_PAGE_FIELDS = """  "drawing_type": "wiring schematic / P&ID / ladder logic / block diagram / layout / other",
  "title": "drawing title or [unclear]",
  "drawing_number": "drawing number from the title block or [unclear]",
  "revision": "latest revision or [unclear]",
  "summary": "2-4 sentence plain-language description of what the drawing shows",
"""

_ITEM_FIELDS = """  "components": [
    {
      "tag": "tag / reference designator",
      "type": "e.g. transmitter, valve, relay, terminal block, PLC module, motor, breaker",
      "description": "description or service text printed on the drawing, otherwise [unclear]",
      "location": "sheet area, panel, or equipment it belongs to, otherwise [unclear]"
    }
  ],
  "connections": [
    {
      "from": "source tag/terminal",
      "to": "destination tag/terminal",
      "line_style": "solid / dashed / dash-dot / other",
      "signal_or_wire_label": "wire number, signal type (4-20mA, 24VDC, discrete...) or [unclear]",
      "wire_count": "integer or [unclear]",
      "notes": "shielding, intrinsic safety, ambiguities"
    }
  ],
  "tables": [
    {
      "name": "what this table is (e.g. Revision table, Bill of materials, I/O list)",
      "headers": ["column 1", "column 2"],
      "rows": [["cell", "cell"]]
    }
  ],
  "text_notes": ["each general note, callout or legend entry, verbatim"]
"""

_RULES = """Rules:
- Only report what is actually drawn; count only wires shown, not spare terminals.
- Copy tags and text exactly as printed, preserving punctuation.
- Every row in a table must have the same number of cells as its headers.
- If something is unclear use [unclear]. Do not invent values. Use empty lists where a section has no content.
"""


def _render_schematic_prompt(page_number: int, total_pages: int) -> str:
    """Whole-page prompt (used when a page is not split into sections)."""
    return (
        "You are an engineering document analyst translating a schematic drawing into tables and text.\n\n"
        f"Analyze page {page_number} of {total_pages} only. If a reference image is provided, use it only for "
        "notation/style context, not as content to transcribe.\n"
        "Transcribe everything meaningful on the page. Work through the drawing region by region so nothing is skipped.\n\n"
        "Return strict JSON in this structure:\n{\n" + _PAGE_FIELDS + _ITEM_FIELDS + "}\n\n" + _RULES
    )


def _render_split_prompt(max_sections: int) -> str:
    return (
        "You are preparing a drawing page for detailed analysis by splitting it into sections.\n"
        "Divide the page into separate sections, one per self-contained circuit or region, for example: a PLC I/O "
        "module with its field wiring, a motor starter or drive circuit, a power distribution block, an instrument "
        "loop or group of related instruments, a legend, a notes block, or a table (title block, revision table, "
        "bill of materials, I/O list).\n"
        "Do not split one circuit across several sections. Sections should not overlap except where a circuit "
        f"genuinely straddles a neighbour. Use at most {max_sections} sections; every circuit, table and note block "
        "on the page must fall inside some section.\n\n"
        "Return strict JSON in this structure:\n{\n" + _PAGE_FIELDS +
        '  "sections": [\n    {\n'
        '      "name": "short label, e.g. \'PLC DI module 3\' or \'Revision table\'",\n'
        '      "description": "one sentence on what this section contains",\n'
        '      "box": [left, top, right, bottom]\n    }\n  ]\n}\n\n'
        "box values are integers from 0 to 1000 measured on the full page image: 0,0 is the top-left corner and "
        "1000,1000 is the bottom-right.\n"
        "If the page is a single circuit with nothing to split, return one section covering the whole page: "
        "[0, 0, 1000, 1000].\n"
    )


def _render_section_prompt(page_number: int, total_pages: int, index: int, count: int, name: str, description: str) -> str:
    return (
        "You are an engineering document analyst translating one section of a schematic drawing into tables and text.\n\n"
        f"This image is a cropped section of page {page_number} of {total_pages}: section {index} of {count}, "
        f'"{name}" ({description}).\n'
        "If a reference image is provided, use it only for notation/style context, not as content to transcribe.\n"
        "Transcribe everything meaningful in this crop. Slivers of neighbouring circuits may show at the edges; "
        "ignore anything cut off that belongs to another section.\n"
        "Where a wire or signal leaves this section (off-page references, continuation arrows, terminal numbers), "
        "record it as a connection and copy its reference label into \"to\" or \"from\" exactly as printed.\n\n"
        "Return strict JSON in this structure:\n{\n" + _ITEM_FIELDS + "}\n\n" + _RULES
    )


def _with_context(prompt: str, context_markdown: Optional[str]) -> str:
    if not context_markdown:
        return prompt
    return f"{prompt}\n\n---\nStructured document context (Docling extraction):\n{context_markdown}"


def _stream_ollama(
    model: str,
    api_host: str,
    api_key: str,
    prompt: str,
    image_paths: List[Path],
    max_tokens: int,
    on_stream: Callable[[str], None],
    num_ctx: Optional[int] = None,
) -> str:
    """Ollama /api/chat with stream=true: newline-delimited JSON, one chunk per token group."""
    headers = {"Content-Type": "application/json"}
    if api_key and api_key != "None":
        headers["Authorization"] = f"Bearer {api_key}"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": "You are a careful engineering document analyst. Return strict JSON for the requested extraction."},
            {"role": "user", "content": prompt, "images": [_load_image_bytes(p) for p in image_paths]},
        ],
        "stream": True,
        "options": {"num_predict": max_tokens, "num_ctx": int(os.getenv("OLLAMA_NUM_CTX") or num_ctx or max_tokens)},
    }
    url = f"{_normalize_host(api_host, 'http://localhost:11434')}/api/chat"
    parts: List[str] = []
    # No read timeout: prompt/image processing can take minutes before the first token.
    with requests.post(url, headers=headers, data=json.dumps(payload), stream=True, timeout=(10, None)) as resp:
        if resp.status_code >= 400:
            raise RuntimeError(f"Ollama request failed: {resp.status_code} {resp.text}")
        for raw in resp.iter_lines(chunk_size=1):
            if not raw:
                continue
            chunk = json.loads(raw)
            if chunk.get("error"):
                raise RuntimeError(f"Ollama error: {chunk['error']}")
            message = chunk.get("message", {})
            if message.get("thinking"):
                on_stream(message["thinking"])
            piece = message.get("content", "")
            if piece:
                parts.append(piece)
                on_stream(piece)
            if chunk.get("done"):
                break
    return "".join(parts).strip()


def _stream_claude(
    model: str,
    api_host: str,
    api_key: str,
    prompt: str,
    image_paths: List[Path],
    max_tokens: int,
    on_stream: Callable[[str], None],
) -> str:
    """Anthropic Messages API with stream=true (server-sent events)."""
    if not api_key or api_key == "None":
        raise RuntimeError("A Claude API key is required.")
    blocks: List[dict] = [{"type": "text", "text": prompt}]
    for path in image_paths:
        blocks.append({"type": "image", "source": {
            "type": "base64", "media_type": "image/png", "data": _load_image_bytes(path)}})
    payload = {
        "model": model,
        "max_tokens": max(1, max_tokens),
        "stream": True,
        "system": "You are an engineering document analyst. Extract the requested information and return valid JSON.",
        "messages": [{"role": "user", "content": blocks}],
    }
    base = (api_host or "https://api.anthropic.com").rstrip("/")
    url = f"{base}/messages" if base.endswith("/v1") else f"{base}/v1/messages"
    headers = {"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"}
    parts: List[str] = []
    with requests.post(url, headers=headers, data=json.dumps(payload), stream=True, timeout=(10, None)) as resp:
        if resp.status_code >= 400:
            raise RuntimeError(f"Claude request failed for model '{model}': {resp.status_code} {resp.text}")
        for raw_line in resp.iter_lines(chunk_size=1):
            raw = raw_line.decode("utf-8") if isinstance(raw_line, bytes) else raw_line
            if not raw.startswith("data:"):
                continue
            event = json.loads(raw[5:].strip())
            if event.get("type") == "error":
                raise RuntimeError(f"Claude stream error: {event.get('error')}")
            delta = event.get("delta", {}) if event.get("type") == "content_block_delta" else {}
            if delta.get("type") == "text_delta":
                parts.append(delta["text"])
                on_stream(delta["text"])
    return "".join(parts).strip()


def _call_model(
    model: str,
    host: str,
    api_key: str,
    prompt: str,
    image_paths: List[Path],
    max_tokens: int,
    log: Callable[[str], None],
    context_markdown: Optional[str],
    on_stream: Optional[Callable[[str], None]] = None,
    num_ctx: Optional[int] = None,
) -> str:
    """max_tokens caps this call's output. num_ctx is Ollama's context window (prompt + image + output);
    it is a separate setting because the output limit can be far larger than the window can hold."""
    listener = on_stream or (lambda _text: None)
    full_prompt = _with_context(prompt, context_markdown)
    if _is_claude_model(model):
        return _stream_claude(model, host, api_key, full_prompt, image_paths, max_tokens, listener)
    if num_ctx:
        max_tokens = min(max_tokens, num_ctx)
    return _stream_ollama(model, host, api_key, full_prompt, image_paths, max_tokens, listener, num_ctx=num_ctx)


SCREEN_PROMPT = """Look at this drawing page.
Question: does this page contain an electrical schematic, or details for electrical components
(wiring diagram, ladder logic, terminal/wiring connections, instrument or I/O lists, panel layouts with electrical devices, electrical equipment schedules)?

Answer with strict JSON only:
{"electrical": true or false, "reason": "one short sentence"}
"""


def _screen_page(
    model: str,
    host: str,
    api_key: str,
    image_path: Path,
    max_tokens: int,
    log: Callable[[str], None],
    on_stream: Optional[Callable[[str], None]],
    num_ctx: Optional[int] = None,
) -> tuple:
    """Quick yes/no pre-check. Returns (is_electrical, reason). Fails open (True) if the answer is unusable."""
    try:
        log("Screening: waiting for first token (the model has to load and read the page image first)...")
        text = _call_model(model, host, api_key, SCREEN_PROMPT, [image_path],
                           min(max_tokens, 512), log, None, on_stream, num_ctx=num_ctx)
        if on_stream:
            on_stream("\n")
    except Exception as exc:
        log(f"Screening failed ({exc}); processing the page anyway")
        return True, "screening failed"
    answer = _extract_json_object(text).get("electrical")
    reason = _cell(_extract_json_object(text).get("reason", ""))
    if isinstance(answer, str):
        answer = answer.strip().lower() in {"true", "yes"}
    if not isinstance(answer, bool):
        log("Screening answer unclear; processing the page anyway")
        return True, "screening unclear"
    return answer, reason


def _is_blank(value: object) -> bool:
    return value is None or str(value).strip() in {"", "[unclear]"}


def _norm(value: object) -> str:
    return " ".join(str(value or "").lower().split())


def _iou(a: List[float], b: List[float]) -> float:
    iw = min(a[2], b[2]) - max(a[0], b[0])
    ih = min(a[3], b[3]) - max(a[1], b[1])
    if iw <= 0 or ih <= 0:
        return 0.0
    inter = iw * ih
    union = (a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - inter
    return inter / union


def _clean_sections(raw: object, max_sections: int, width: int, height: int) -> List[dict]:
    """Validate model-supplied sections into boxes on a 0-1000 scale, tolerating common scale mistakes."""
    if not isinstance(raw, list):
        return []
    sections: List[dict] = []
    for i, item in enumerate(raw, start=1):
        if not isinstance(item, dict):
            continue
        try:
            x1, y1, x2, y2 = (float(v) for v in (item.get("box") or item.get("bbox") or item.get("bounding_box")))
        except (TypeError, ValueError):
            continue
        top = max(x1, y1, x2, y2)
        if top <= 1.0:  # fractions of the page
            x1, y1, x2, y2 = (v * 1000 for v in (x1, y1, x2, y2))
        elif top > 1000:  # pixels of the rendered image
            x1, x2 = x1 * 1000 / width, x2 * 1000 / width
            y1, y2 = y1 * 1000 / height, y2 * 1000 / height
        x1, x2 = sorted((min(max(x1, 0), 1000), min(max(x2, 0), 1000)))
        y1, y2 = sorted((min(max(y1, 0), 1000), min(max(y2, 0), 1000)))
        if x2 - x1 < 30 or y2 - y1 < 30:  # under 3% of the page: not a real section
            continue
        box = [x1, y1, x2, y2]
        if any(_iou(box, kept["box"]) > 0.7 for kept in sections):
            continue
        sections.append({
            "name": _cell(item.get("name")).strip() or f"Section {i}",
            "description": _cell(item.get("description")).strip() or "[unclear]",
            "box": [round(v) for v in box],
        })
        if len(sections) >= max_sections:
            break
    return sections


def _coverage(sections: List[dict]) -> float:
    """Fraction of the page covered by the union of all section boxes (50x50 grid)."""
    n = 50
    covered = set()
    for s in sections:
        x1, y1, x2, y2 = (int(v * n / 1000) for v in s["box"])
        covered.update((x, y) for x in range(x1, max(x2, x1 + 1)) for y in range(y1, max(y2, y1 + 1)))
    return len(covered) / (n * n)


def _crop_section(image_path: Path, box: List[int], out_path: Path, pad: float = 0.025) -> Path:
    """Crop a 0-1000 box (plus a small margin so edge text isn't clipped) out of the rendered page."""
    with Image.open(image_path) as img:
        w, h = img.size
        left = max(int((box[0] / 1000 - pad) * w), 0)
        top = max(int((box[1] / 1000 - pad) * h), 0)
        right = min(int((box[2] / 1000 + pad) * w), w)
        bottom = min(int((box[3] / 1000 + pad) * h), h)
        img.crop((left, top, right, bottom)).save(out_path)
    return out_path


def _save_overlay(image_path: Path, sections: List[dict], out_path: Path) -> None:
    """Draw the section boxes on the page so the model's splitting can be checked by eye."""
    colors = ["#e6194b", "#3cb44b", "#4363d8", "#f58231", "#911eb4", "#008080", "#f032e6", "#9a6324"]
    with Image.open(image_path) as src:
        img = src.convert("RGB")
    w, h = img.size
    draw = ImageDraw.Draw(img)
    line = max(w // 400, 3)
    for i, s in enumerate(sections):
        color = colors[i % len(colors)]
        x1, y1, x2, y2 = s["box"][0] * w / 1000, s["box"][1] * h / 1000, s["box"][2] * w / 1000, s["box"][3] * h / 1000
        draw.rectangle([x1, y1, x2, y2], outline=color, width=line)
        draw.text((x1 + line * 2, y1 + line * 2), f"{i + 1}. {s['name']}", fill=color, font_size=max(w // 70, 14))
    if w > 2400:
        img = img.resize((2400, int(h * 2400 / w)))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path)


def _query_json(
    label: str,
    model: str,
    host: str,
    api_key: str,
    prompt: str,
    images: List[Path],
    max_tokens: int,
    log: Callable[[str], None],
    ocr_text: Optional[str],
    on_stream: Optional[Callable[[str], None]],
    num_ctx: Optional[int] = None,
) -> dict:
    """Ask the model for JSON, retrying (with back-off) on errors or unparsable output."""
    prompt = with_reference_note(prompt, len(images) > 1)  # more than one image means a reference image leads
    for attempt in range(1, 4):
        try:
            if on_stream:
                log(f"{label} attempt {attempt}: waiting for first token (image processing can take a while)...")
            text = _call_model(model, host, api_key, prompt, images, max_tokens, log, ocr_text, on_stream, num_ctx)
            if on_stream:
                on_stream("\n")
        except Exception as exc:
            log(f"{label} attempt {attempt} failed: {exc}")
            text = ""
            if attempt < 3:
                overloaded = "529" in str(exc) or "overload" in str(exc).lower()
                time.sleep((30 if overloaded else 5) * attempt)
        parsed = _extract_json_object(text) if text.strip() else {}
        if parsed:
            return parsed
        log(f"{label} attempt {attempt} produced no parsable JSON")
    return {}


def _merge_into(existing: dict, new: dict, section: str) -> None:
    """Fill blanks in an already-collected item from a duplicate seen in another section."""
    for key, value in new.items():
        if _is_blank(existing.get(key)) and not _is_blank(value):
            existing[key] = value
    seen = [s.strip() for s in _cell(existing.get("section")).split(",") if s.strip()]
    if section and section not in seen:
        existing["section"] = ", ".join(seen + [section])


def _merge_section_results(meta: dict, parts: List[tuple]) -> dict:
    """Combine per-section results into one page result, de-duplicating items that overlapping sections both saw."""
    components: dict = {}
    connections: dict = {}
    tables: dict = {}
    notes: List[str] = []
    for section, parsed in parts:
        name = section["name"]
        for item in parsed.get("components", []):
            if not isinstance(item, dict):
                continue
            key = _norm(item.get("tag"))
            if _is_blank(key):
                key = f"__untagged_{len(components)}"
            if key in components:
                _merge_into(components[key], item, name)
            else:
                components[key] = {**item, "section": name}
        for item in parsed.get("connections", []):
            if not isinstance(item, dict):
                continue
            key = (_norm(item.get("from")), _norm(item.get("to")), _norm(item.get("line_style")))
            if _is_blank(key[0]) and _is_blank(key[1]):
                key = ("__unlinked", str(len(connections)), "")
            if key in connections:
                _merge_into(connections[key], item, name)
            else:
                connections[key] = {**item, "section": name}
        for table in parsed.get("tables", []):
            if not isinstance(table, dict):
                continue
            key = json.dumps([table.get("name"), table.get("headers"), table.get("rows")], sort_keys=True, default=str)
            tables.setdefault(key, {**table, "section": name})
        for note in parsed.get("text_notes", []):
            text = _cell(note).strip()
            if text and text not in notes:
                notes.append(text)
    merged = {k: meta.get(k, "[unclear]") for k in ("drawing_type", "title", "drawing_number", "revision", "summary")}
    merged.update({
        "components": list(components.values()),
        "connections": list(connections.values()),
        "tables": list(tables.values()),
        "text_notes": notes,
        "sections": [dict(s) for s, _ in parts],
    })
    return merged


def _translate_sections(
    index: int,
    total: int,
    image_path: Path,
    work_dir: Path,
    overlay_path: Optional[Path],
    model: str,
    api_host: str,
    api_key: str,
    max_tokens: int,
    max_sections: int,
    reference_image: Optional[Path],
    ocr_text: Optional[str],
    log: Callable[[str], None],
    on_stream: Optional[Callable[[str], None]],
    num_ctx: Optional[int] = None,
) -> Optional[dict]:
    """Split a page into circuit sections and translate each. Returns None if the page can't usefully be split."""
    log(f"Page {index}: asking the model to split the page into circuit sections...")
    split = _query_json(f"Page {index} section split", model, api_host, api_key,
                        _render_split_prompt(max_sections), [image_path], min(max_tokens, 2048),
                        log, None, on_stream, num_ctx=num_ctx)
    with Image.open(image_path) as img:
        width, height = img.size
    sections = _clean_sections(split.get("sections"), max_sections, width, height)
    if len(sections) < 2:
        log(f"Page {index}: {len(sections)} usable section(s) returned; analyzing the whole page instead")
        return None

    coverage = _coverage(sections)
    log(f"Page {index}: {len(sections)} sections covering {coverage:.0%} of the page")
    if coverage < 0.5:
        log(f"Page {index}: WARNING sections cover under half the page; content outside them will be missed "
            "(check the overlay image)")
    if overlay_path:
        _save_overlay(image_path, sections, overlay_path)
        log(f"Page {index}: section overlay saved to {overlay_path}")

    parts: List[tuple] = []
    failed: List[str] = []
    for n, section in enumerate(sections, start=1):
        label = f"Page {index} section {n}/{len(sections)} '{section['name']}'"
        log(f"--- {label} ---")
        crop = _crop_section(image_path, section["box"], work_dir / f"page_{index:03d}_section_{n:02d}.png")
        prompt = _render_section_prompt(index, total, n, len(sections), section["name"], section["description"])
        images = ([reference_image] if reference_image else []) + [crop]
        parsed = _query_json(label, model, api_host, api_key, prompt, images, max_tokens, log, ocr_text, on_stream,
                             num_ctx)
        if not parsed:
            failed.append(section["name"])
            continue
        parts.append((section, parsed))
        log(f"{label}: {len(parsed.get('components', []))} components, "
            f"{len(parsed.get('connections', []))} connections, {len(parsed.get('tables', []))} tables")

    if not parts:
        log(f"Page {index}: every section failed; analyzing the whole page instead")
        return None
    merged = _merge_section_results(split, parts)
    if failed:
        merged["notes"] = f"No parsable output for section(s): {', '.join(failed)}"
        log(f"Page {index}: WARNING sections with no output: {', '.join(failed)}")
    return merged


def translate_pdf(
    pdf_path: Path,
    model: str,
    api_host: str,
    api_key: str,
    max_tokens: int = 8192,
    dpi: int = 300,
    pages: Optional[List[int]] = None,
    reference_image: Optional[Path] = None,
    use_ocr: bool = True,
    ocr_scale: float = 4.0,
    log: Callable[[str], None] = print,
    on_stream: Optional[Callable[[str], None]] = None,
    use_sections: bool = True,
    max_sections: int = 12,
    overlay_dir: Optional[Path] = None,
    num_ctx: int = 16384,
) -> dict:
    """Send each selected page to the model and return {'source': ..., 'pages': [...]}.

    With use_sections, each page is first split into circuit sections which are cropped and translated
    one at a time, then merged; pages that can't be split fall back to a single whole-page pass.
    """
    work_dir = Path(tempfile.mkdtemp(prefix="schematic_pages_"))
    try:
        image_paths = render_pdf_pages_to_images(pdf_path, dpi=dpi, output_dir=work_dir)
        total = len(image_paths)
        selected = pages or list(range(1, total + 1))
        log(f"Translating {len(selected)} of {total} page(s) from {pdf_path.name} using '{model}'"
            f"{' (section mode)' if use_sections else ''}")

        if not _is_claude_model(model):
            if num_ctx < 8192:
                log(f"WARNING: context window ({num_ctx}) is small. A page image alone needs roughly 4-5k tokens, "
                    "so use 16384 or more or requests will fail with 'exceeds the available context size'.")
            elif num_ctx > 65536:
                log(f"WARNING: context window ({num_ctx}) is very large. Ollama must allocate memory for all of it, "
                    "which can make the model fail to load or crawl. 16384-32768 is plenty for one page or section.")
            log(f"Ollama context window {num_ctx}, max output {min(max_tokens, num_ctx)} tokens per call")

        # Pre-step: full-page OCR of every selected page, done once up front (cached on disk).
        ocr_pages = preprocess_pdf_ocr(pdf_path, pages=selected, images_scale=ocr_scale, log_callback=log) if use_ocr else {}

        results: List[dict] = []
        for index, image_path in enumerate(image_paths, start=1):
            if index not in selected:
                continue
            log(f"--- Page {index}/{total} ---")

            log(f"Page {index}: checking for electrical schematic content...")
            is_electrical, reason = _screen_page(model, api_host, api_key, image_path, max_tokens, log, on_stream, num_ctx)
            if not is_electrical:
                log(f"Page {index}: no electrical content ({reason}); skipping")
                results.append({"page": index, "title": "[skipped]",
                                "summary": f"Skipped - no electrical schematic or component details: {reason}"})
                continue
            log(f"Page {index}: electrical content found ({reason}); translating")

            ocr_text = ocr_pages.get(index)
            if use_ocr and not ocr_text:
                log("OCR returned nothing; continuing with the image only")

            parsed: Optional[dict] = None
            if use_sections:
                overlay = overlay_dir / f"page_{index:03d}_sections.png" if overlay_dir else None
                parsed = _translate_sections(
                    index, total, image_path, work_dir, overlay, model, api_host, api_key, max_tokens,
                    max_sections, reference_image, ocr_text, log, on_stream, num_ctx,
                )

            if parsed is None:
                prompt = _render_schematic_prompt(index, total)
                images = ([reference_image] if reference_image else []) + [image_path]
                parsed = _query_json(f"Page {index}", model, api_host, api_key, prompt, images,
                                     max_tokens, log, ocr_text, on_stream, num_ctx)
            if not parsed:
                parsed = {"summary": "[unclear]", "notes": "Model returned no parsable output for this page."}
            results.append({"page": index, **parsed})
            log(
                f"Page {index}: {len(parsed.get('components', []))} components, "
                f"{len(parsed.get('connections', []))} connections, "
                f"{len(parsed.get('tables', []))} tables"
            )
        return {"source": pdf_path.name, "model": model, "pages": results}
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def _cell(value: object) -> str:
    """Flatten any value into a single-line table-safe string."""
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return "; ".join(_cell(v) for v in value)
    return str(value)


def _md_table(headers: List[str], rows: List[List[str]]) -> str:
    if not headers:
        return ""
    width = len(headers)
    esc = lambda s: _cell(s).replace("|", "\\|").replace("\n", " ")
    lines = ["| " + " | ".join(esc(h) for h in headers) + " |", "|" + " --- |" * width]
    for row in rows:
        row = (list(row) + [""] * width)[:width]
        lines.append("| " + " | ".join(esc(c) for c in row) + " |")
    return "\n".join(lines)


def _rows_from(items: List[dict], keys: List[str]) -> List[List[str]]:
    return [[_cell(item.get(k, "")) for k in keys] for item in items if isinstance(item, dict)]


COMPONENT_KEYS = ["tag", "type", "description", "location", "section"]
CONNECTION_KEYS = ["from", "to", "line_style", "signal_or_wire_label", "wire_count", "notes", "section"]


def to_markdown(data: dict) -> str:
    out = [f"# Schematic translation: {data.get('source', '')}", ""]
    for page in data.get("pages", []):
        out.append(f"## Page {page.get('page')}: {page.get('title', '[unclear]')}")
        out.append(
            f"*{page.get('drawing_type', '[unclear]')} | Drawing no. {page.get('drawing_number', '[unclear]')}"
            f" | Rev {page.get('revision', '[unclear]')}*"
        )
        out += ["", _cell(page.get("summary", "")), ""]
        if page.get("notes"):
            out += [f"> {_cell(page['notes'])}", ""]
        sections = [x for x in page.get("sections", []) if isinstance(x, dict)]
        if sections:
            out += ["### Sections", _md_table(["name", "description", "box"], _rows_from(sections, ["name", "description", "box"])), ""]
        comps = [c for c in page.get("components", []) if isinstance(c, dict)]
        if comps:
            out += ["### Components", _md_table(COMPONENT_KEYS, _rows_from(comps, COMPONENT_KEYS)), ""]
        conns = [c for c in page.get("connections", []) if isinstance(c, dict)]
        if conns:
            out += ["### Connections", _md_table(CONNECTION_KEYS, _rows_from(conns, CONNECTION_KEYS)), ""]
        for table in page.get("tables", []):
            if isinstance(table, dict):
                where = f" (section: {table['section']})" if table.get("section") else ""
                out += [f"### {table.get('name', 'Table')}{where}", _md_table(table.get("headers", []), table.get("rows", [])), ""]
        notes = page.get("text_notes", [])
        if notes:
            out += ["### Notes"] + [f"- {_cell(n)}" for n in notes] + [""]
    return "\n".join(out)


def write_excel(data: dict, excel_path: Path) -> None:
    if openpyxl is None:
        raise RuntimeError("openpyxl is required for Excel export. Install it with: pip install openpyxl")
    wb = openpyxl.Workbook()
    summary = wb.active
    summary.title = "Pages"
    summary.append(["Page", "Title", "Drawing Type", "Drawing Number", "Revision", "Summary"])
    comps = wb.create_sheet("Components")
    comps.append(["Page"] + [k.title() for k in COMPONENT_KEYS])
    conns = wb.create_sheet("Connections")
    conns.append(["Page", "From", "To", "Line Style", "Signal / Wire Label", "Wire Count", "Notes", "Section"])
    tables = wb.create_sheet("Drawing Tables")
    notes = wb.create_sheet("Notes")
    notes.append(["Page", "Note"])
    sections_ws = wb.create_sheet("Sections")
    sections_ws.append(["Page", "Name", "Description", "Box [left, top, right, bottom] (0-1000)"])

    for page in data.get("pages", []):
        n = page.get("page", "")
        summary.append([n] + [_cell(page.get(k, "")) for k in
                              ["title", "drawing_type", "drawing_number", "revision", "summary"]])
        for row in _rows_from(page.get("components", []), COMPONENT_KEYS):
            comps.append([n] + row)
        for row in _rows_from(page.get("connections", []), CONNECTION_KEYS):
            conns.append([n] + row)
        for table in page.get("tables", []):
            if not isinstance(table, dict):
                continue
            tables.append([f"Page {n}", _cell(table.get("name", "Table")), _cell(table.get("section", ""))])
            tables.append([""] + [_cell(h) for h in table.get("headers", [])])
            for row in table.get("rows", []):
                tables.append([""] + [_cell(c) for c in row])
            tables.append([])
        for note in page.get("text_notes", []):
            notes.append([n, _cell(note)])
        for sec in page.get("sections", []):
            if isinstance(sec, dict):
                sections_ws.append([n] + _rows_from([sec], ["name", "description", "box"])[0])

    for ws in wb.worksheets:
        for col in ws.columns:
            longest = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(longest + 4, 60)
    wb.save(str(excel_path))


GUI_INPUTS_FILE = Path(__file__).with_name("schematic_translator_gui_inputs.txt")


def run_job(
    pdf: Path,
    model: str,
    api_host: str,
    api_key: str,
    max_tokens: int,
    dpi: int,
    pages_text: str,
    reference_image: Optional[Path],
    use_ocr: bool,
    ocr_scale: float,
    output: Optional[Path],
    log: Callable[[str], None],
    on_stream: Optional[Callable[[str], None]] = None,
    use_sections: bool = True,
    max_sections: int = 12,
    num_ctx: int = 16384,
) -> Path:
    """Shared by CLI and GUI: translate, write .json/.md/.xlsx, return the output base path."""
    import fitz  # PyMuPDF; Instrument_List_Transcriber already requires it

    with fitz.open(str(pdf)) as doc:
        page_count = len(doc)
    pages = parse_page_selection(pages_text, page_count) if pages_text.strip() else None
    if pages is not None and not pages:
        raise ValueError(f"No valid pages in '{pages_text}' (PDF has {page_count} pages)")

    if _is_claude_model(model) and "anthropic" not in api_host:
        api_host = "https://api.anthropic.com"

    base = (output or pdf.with_name(pdf.stem + " - translated")).with_suffix("")
    data = translate_pdf(
        pdf, model, api_host, api_key or "None", max_tokens, dpi, pages,
        reference_image, use_ocr, ocr_scale, log, on_stream,
        use_sections=use_sections, max_sections=max_sections,
        overlay_dir=base.parent / f"{base.name} - sections", num_ctx=num_ctx,
    )

    base.with_suffix(".json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    base.with_suffix(".md").write_text(to_markdown(data), encoding="utf-8")
    write_excel(data, base.with_suffix(".xlsx"))
    log(f"Wrote {base.with_suffix('.json')}, .md and .xlsx")
    return base


def launch_gui() -> int:
    """Settings form on the left, live run log on the right."""
    try:
        saved = json.loads(GUI_INPUTS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        saved = {}

    root = tk.Tk()
    root.title("Schematic To Table Translator")
    width, height = 1500, 850
    root.geometry(f"{width}x{height}+{max((root.winfo_screenwidth() - width) // 2, 0)}+"
                  f"{max((root.winfo_screenheight() - height) // 2, 0)}")
    root.minsize(1000, 600)

    defaults = {
        "pdf": "", "model": os.getenv("AI_GATEWAY_MODEL", "qwen3-vl:4b-instruct"),
        "api_host": os.getenv("AI_GATEWAY_HOST", "http://localhost:11434"),
        "api_key": os.getenv("AI_GATEWAY_API_KEY", ""), "max_tokens": "8192", "dpi": "300",
        "pages": "", "reference_image": "", "output": "", "ocr": True, "ocr_scale": "4.0",
        "sections": True, "max_sections": "12", "num_ctx": "16384",
    }
    vars_: dict = {}
    for key, default in defaults.items():
        value = saved.get(key, default)
        vars_[key] = tk.BooleanVar(value=bool(value)) if isinstance(default, bool) else tk.StringVar(value=str(value))

    paned = ttk.Panedwindow(root, orient=tk.HORIZONTAL)
    paned.pack(fill="both", expand=True)
    left, right = ttk.Frame(paned, padding=12), ttk.Frame(paned, padding=12)
    paned.add(left, weight=1)
    paned.add(right, weight=2)
    left.columnconfigure(1, weight=1)

    log_text = tk.Text(right, wrap="word", state="disabled")
    log_scroll = ttk.Scrollbar(right, orient="vertical", command=log_text.yview)
    log_text.configure(yscrollcommand=log_scroll.set)
    ttk.Label(right, text="Run Log", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
    log_btns = ttk.Frame(right)
    log_btns.pack(side="bottom", anchor="w", pady=(8, 0))
    log_scroll.pack(side="right", fill="y")
    log_text.pack(side="left", fill="both", expand=True)

    # Worker thread -> UI thread: ("line", text) for timestamped log lines, ("chunk", text) for streamed tokens.
    log_queue: "queue.Queue[tuple[str, str]]" = queue.Queue()
    status = tk.StringVar(value="Ready")

    def append_log(message: str) -> None:
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        log_text.configure(state="normal")
        log_text.insert(tk.END, f"[{stamp}] {message}\n")
        log_text.configure(state="disabled")
        log_text.see(tk.END)

    def append_chunk(text: str) -> None:
        log_text.configure(state="normal")
        log_text.insert(tk.END, text)
        log_text.configure(state="disabled")
        log_text.see(tk.END)

    def pump_log() -> None:
        try:
            while True:
                kind, text = log_queue.get_nowait()
                (append_chunk if kind == "chunk" else append_log)(text)
        except queue.Empty:
            pass
        root.after(100, pump_log)

    def save_log() -> None:
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt")])
        if path:
            Path(path).write_text(log_text.get("1.0", tk.END), encoding="utf-8")

    def clear_log() -> None:
        log_text.configure(state="normal")
        log_text.delete("1.0", tk.END)
        log_text.configure(state="disabled")

    ttk.Button(log_btns, text="Save Log", command=save_log).pack(side="left", padx=(0, 6))
    ttk.Button(log_btns, text="Clear Log", command=clear_log).pack(side="left")

    row = [0]

    def add_row(label: str, key: str, browse: Optional[str] = None, filetypes=None, show: str = "") -> None:
        ttk.Label(left, text=label).grid(row=row[0], column=0, sticky="w", pady=4, padx=(0, 8))
        ttk.Entry(left, textvariable=vars_[key], show=show).grid(row=row[0], column=1, sticky="ew", pady=4)
        if browse:
            def pick() -> None:
                types = filetypes or [("All", "*.*")]
                if browse == "save":
                    path = filedialog.asksaveasfilename(title=label, filetypes=types)
                else:
                    path = filedialog.askopenfilename(title=label, filetypes=types)
                if path:
                    vars_[key].set(path)
            ttk.Button(left, text="Browse", command=pick).grid(row=row[0], column=2, padx=(6, 0))
        row[0] += 1

    def section(title: str) -> None:
        ttk.Label(left, text=title, font=("Segoe UI", 10, "bold")).grid(
            row=row[0], column=0, columnspan=3, sticky="w", pady=(12, 2))
        row[0] += 1

    section("Input")
    add_row("Drawing PDF", "pdf", "open", [("PDF", "*.pdf")])
    add_row("Pages (blank = all)", "pages")
    add_row("Reference image (optional)", "reference_image", "open",
            [("Images", "*.png *.jpg *.jpeg"), ("All", "*.*")])

    section("Model")
    ttk.Label(left, text="Model").grid(row=row[0], column=0, sticky="w", pady=4)
    model_box = ttk.Combobox(left, textvariable=vars_["model"])
    model_box.grid(row=row[0], column=1, sticky="ew", pady=4)

    def refresh_models() -> None:
        names = _get_local_ollama_models(vars_["api_host"].get().strip() or "http://localhost:11434")
        names += _get_claude_models(vars_["api_key"].get().strip())
        model_box["values"] = names
        append_log(f"Found {len(names)} model(s)" if names else
                   "No models found (is Ollama running / is the Claude API key set?)")

    ttk.Button(left, text="Refresh", command=refresh_models).grid(row=row[0], column=2, padx=(6, 0))
    row[0] += 1
    add_row("API host", "api_host")
    add_row("API key", "api_key", show="*")
    add_row("Max output tokens", "max_tokens")
    add_row("Context window (Ollama num_ctx)", "num_ctx")

    section("Page sectioning")
    ttk.Checkbutton(left, text="Split each page into circuit sections first (more accurate, slower)",
                    variable=vars_["sections"]).grid(row=row[0], column=0, columnspan=3, sticky="w", pady=4)
    row[0] += 1
    add_row("Max sections per page", "max_sections")

    section("Rendering / OCR")
    add_row("Render DPI", "dpi")
    ttk.Checkbutton(left, text="Full-page OCR pre-step (Docling/RapidOCR, cached)",
                    variable=vars_["ocr"]).grid(row=row[0], column=0, columnspan=3, sticky="w", pady=4)
    row[0] += 1
    add_row("OCR image scale", "ocr_scale")

    section("Output")
    add_row("Output base path (blank = next to PDF)", "output", "save", [("All", "*.*")])

    run_btn = ttk.Button(left, text="Translate")
    run_btn.grid(row=row[0], column=0, columnspan=3, sticky="ew", pady=(18, 4))
    ttk.Label(left, textvariable=status).grid(row=row[0] + 1, column=0, columnspan=3, sticky="w")

    def start() -> None:
        try:
            pdf = Path(vars_["pdf"].get().strip())
            if not pdf.is_file():
                raise ValueError("Choose an existing PDF.")
            ref_text = vars_["reference_image"].get().strip()
            ref = Path(ref_text) if ref_text else None
            if ref and not ref.is_file():
                raise ValueError("Reference image not found.")
            model = vars_["model"].get().strip()
            if not model:
                raise ValueError("Enter a model name.")
            max_tokens = int(vars_["max_tokens"].get())
            dpi = int(vars_["dpi"].get())
            ocr_scale = float(vars_["ocr_scale"].get())
            if min(max_tokens, dpi) < 1 or ocr_scale <= 0:
                raise ValueError("Max tokens, DPI and OCR scale must be positive.")
            num_ctx = int(vars_["num_ctx"].get())
            if num_ctx < 2048:
                raise ValueError("Context window must be at least 2048 (a page image alone needs ~4-5k).")
            max_sections = int(vars_["max_sections"].get())
            if max_sections < 2:
                raise ValueError("Max sections per page must be at least 2.")
            out_text = vars_["output"].get().strip()
            job = (pdf, model, vars_["api_host"].get().strip(), vars_["api_key"].get().strip(), max_tokens,
                   dpi, vars_["pages"].get(), ref, vars_["ocr"].get(), ocr_scale,
                   Path(out_text) if out_text else None)
        except ValueError as exc:
            messagebox.showerror("Invalid input", str(exc))
            return

        use_sections = vars_["sections"].get()
        GUI_INPUTS_FILE.write_text(json.dumps({k: v.get() for k, v in vars_.items()}, indent=2), encoding="utf-8")
        run_btn.configure(state="disabled")
        status.set("Running...")

        def worker() -> None:
            try:
                base = run_job(*job, log=lambda m: log_queue.put(("line", m)),
                               on_stream=lambda t: log_queue.put(("chunk", t)),
                               use_sections=use_sections, max_sections=max_sections,
                               num_ctx=num_ctx)
                log_queue.put(("line", f"Done: {base.with_suffix('.xlsx')}"))
                root.after(0, lambda: status.set("Finished"))
            except Exception as exc:
                import traceback
                log_queue.put(("line", f"ERROR: {exc}\n{traceback.format_exc()}"))
                root.after(0, lambda: status.set("Failed - see log"))
            finally:
                root.after(0, lambda: run_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    run_btn.configure(command=start)
    pump_log()
    root.mainloop()
    return 0


def main() -> int:
    if len(sys.argv) == 1:
        return launch_gui()

    parser = argparse.ArgumentParser(
        description="Translate PDF schematics into tables and text. Run with no arguments for the GUI.")
    parser.add_argument("pdf", type=Path, help="Path to the drawing PDF")
    parser.add_argument("--model", default=os.getenv("AI_GATEWAY_MODEL", "qwen3-vl:4b-instruct"))
    parser.add_argument("--api-host", default=os.getenv("AI_GATEWAY_HOST", "http://localhost:11434"))
    parser.add_argument("--api-key", default=os.getenv("AI_GATEWAY_API_KEY", "None"))
    parser.add_argument("--max-tokens", type=int, default=8192, help="Maximum output tokens per call")
    parser.add_argument("--num-ctx", type=int, default=16384,
                        help="Ollama context window (prompt + image + output); keep it modest, e.g. 16384-32768")
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--pages", default="", help="e.g. '1,3-5'; default is all pages")
    parser.add_argument("--reference-image", type=Path, help="Optional notation reference image")
    parser.add_argument("--no-ocr", action="store_true",
                        help="Skip the full-page OCR pre-step (on by default; results are cached in ocr_cache/)")
    parser.add_argument("--ocr-scale", type=float, default=4.0)
    parser.add_argument("--no-sections", action="store_true",
                        help="Translate each page in one pass instead of splitting it into circuit sections first")
    parser.add_argument("--max-sections", type=int, default=12, help="Maximum sections per page (default 12)")
    parser.add_argument("--output", type=Path, help="Output base path (default: next to the PDF)")
    args = parser.parse_args()

    if not args.pdf.is_file():
        print(f"PDF not found: {args.pdf}", file=sys.stderr)
        return 1
    if args.reference_image and not args.reference_image.is_file():
        print(f"Reference image not found: {args.reference_image}", file=sys.stderr)
        return 1

    run_job(args.pdf, args.model, args.api_host, args.api_key, args.max_tokens, args.dpi, args.pages,
            args.reference_image, not args.no_ocr, args.ocr_scale, args.output, print,
            on_stream=lambda t: print(t, end="", flush=True),
            use_sections=not args.no_sections, max_sections=args.max_sections,
            num_ctx=args.num_ctx)
    return 0


if __name__ == "__main__":
    sys.exit(main())
