print("Initializing")
import argparse
print("import argparse")
import base64
print("import base64")
import json
print("import json")
import os
print("import os")
import re
print("import re")
import sys
print("import sys")
import tempfile
print("import tempfile")
import threading
print("import threading")
import datetime
print("import datetime")
import logging
print("import logging")
from pathlib import Path
print("from pathlib import Path")
from typing import Callable, List, Optional
print("from typing import Callable, List, Optional")
import tkinter as tk
print("import tkinter as tk")
from tkinter import filedialog, messagebox, ttk
print("from tkinter import filedialog, messagebox, ttk")

import requests

from Pdf_OCR_Preprocess import preprocess_pdf_ocr
from Claude_Code_Backend import is_claude_code_model, send_request as _claude_code_send
print("import requests")

try:
    import fitz  # PyMuPDF
    print("import fitz")
except ImportError as exc:  # pragma: no cover - runtime dependency guard
    fitz = None
    _FITZ_IMPORT_ERROR = exc
else:
    _FITZ_IMPORT_ERROR = None

try:
    import openpyxl
    print("import openpyxl")
except ImportError:  # pragma: no cover
    openpyxl = None

try:
    from docling.document_converter import DocumentConverter, PdfFormatOption
    print("from docling.document_converter import DocumentConverter, PdfFormatOption")
    from docling.datamodel.base_models import InputFormat
    print("from docling.datamodel.base_models import InputFormat")
    from docling.pipeline.vlm_pipeline import VlmPipeline
    print("from docling.pipeline.vlm_pipeline import VlmPipeline")
    from docling.datamodel.pipeline_options import (
        VlmPipelineOptions,
        PdfPipelineOptions,
        TesseractOcrOptions,
        RapidOcrOptions,
    )
    print("from docling.datamodel.pipeline_options import (VlmPipelineOptions, PdfPipelineOptions, TesseractOcrOptions, RapidOcrOptions)")
    from docling.datamodel import vlm_model_specs
    print("from docling.datamodel import vlm_model_specs")
    from docling_core.types.doc.base import ImageRefMode
    print("from docling_core.types.doc.base import ImageRefMode")
    DOCLING_AVAILABLE = True
except ImportError:
    DOCLING_AVAILABLE = False


class CallbackLoggingHandler(logging.Handler):
    """Logging handler that sends log records to a callback function."""

    def __init__(self, callback: Optional[Callable[[str], None]] = None, level=logging.INFO):
        super().__init__(level)
        self.callback = callback

    def emit(self, record: logging.LogRecord) -> None:
        if self.callback:
            try:
                msg = self.format(record)
                self.callback(f"[{record.name}] {msg}")
            except Exception:
                pass


def _extract_page_markdown_from_pdf(
    pdf_path: Path,
    page_numbers: Optional[List[int]] = None,
    do_ocr: bool = True,
    do_table_structure: bool = True,
    do_picture_description: bool = False,
    images_scale: float = 1.0,
    log_callback: Optional[Callable[[str], None]] = None,
) -> Optional[str]:
    """Extract markdown for a single page from a PDF using Docling with configurable options."""
    if not DOCLING_AVAILABLE:
        if log_callback:
            log_callback("[DOCLING] ERROR: Docling not available (not installed)")
        return None

    if fitz is None:
        if log_callback:
            log_callback("[DOCLING] ERROR: PyMuPDF required for page extraction")
        return None

    def _docling_log(msg: str) -> None:
        if log_callback:
            log_callback(f"[DOCLING] {msg}")

    handler = None
    handlers = []

    try:
        if not page_numbers or len(page_numbers) == 0:
            _docling_log("No pages specified for extraction")
            return None

        page_num = page_numbers[0]  # Process only first page in list
        _docling_log(f"Extracting page {page_num} from {pdf_path.name}")
        _docling_log(f"  OCR: {do_ocr}, Tables: {do_table_structure}, Pictures: {do_picture_description}, Scale: {images_scale}x")

        if log_callback:
            handler = CallbackLoggingHandler(log_callback, level=logging.DEBUG)
            formatter = logging.Formatter("%(levelname)s: %(message)s")
            handler.setFormatter(formatter)
            for logger_name in ["docling", "docling_core", "docling.backend", "PIL", "pydantic"]:
                logger = logging.getLogger(logger_name)
                logger.setLevel(logging.DEBUG)
                logger.propagate = True
                logger.addHandler(handler)
                handlers.append((logger, handler))

        _docling_log("Extracting page from PDF...")
        doc_pdf = fitz.open(str(pdf_path))
        if page_num < 1 or page_num > len(doc_pdf):
            _docling_log(f"ERROR: Page {page_num} out of range (1-{len(doc_pdf)})")
            doc_pdf.close()
            return None

        temp_pdf_path = Path(tempfile.mktemp(suffix=".pdf"))
        try:
            # Rasterize the page into an image-only PDF so any embedded text layer is
            # discarded and Docling's OCR is the sole source of text.
            source_page = doc_pdf[page_num - 1]
            pixmap = source_page.get_pixmap(dpi=300)
            output_pdf = fitz.open()
            new_page = output_pdf.new_page(width=source_page.rect.width, height=source_page.rect.height)
            new_page.insert_image(new_page.rect, pixmap=pixmap)
            _docling_log("Rasterized page to image-only PDF (embedded text layer removed)")
            output_pdf.save(str(temp_pdf_path))
            output_pdf.close()
            doc_pdf.close()

            # Use RapidOCR with aggressive settings for maximum P&ID detail extraction
            rapidocr_options = RapidOcrOptions(
                force_full_page_ocr=True,
                bitmap_area_threshold=0.001,
            )

            pdf_pipeline_options = PdfPipelineOptions(
                ocr_options=rapidocr_options,
                do_table_structure=do_table_structure,
                do_picture_description=do_picture_description,
                images_scale=images_scale,
            )

            _docling_log(f"Docling OCR Pipeline Configuration (RapidOCR):")
            _docling_log(f"  Engine: RapidOCR (ONNX/OpenVINO, optimized for P&IDs)")
            _docling_log(f"  Full-page OCR: True (extracts all text, including small tags)")
            _docling_log(f"  Bitmap threshold: 0.001 (maximum sensitivity for instrument tags)")
            _docling_log(f"  Table structure: {do_table_structure}")
            _docling_log(f"  Picture description: {do_picture_description}")
            _docling_log(f"  Images scale: {images_scale}x (for fine schematic details)")

            _docling_log("Initializing DocumentConverter with RapidOCR (StandardPdfPipeline)...")
            converter = DocumentConverter(
                format_options={
                    InputFormat.PDF: PdfFormatOption(
                        pipeline_options=pdf_pipeline_options,
                    ),
                }
            )

            _docling_log("Converting page with Docling...")
            conv_result = converter.convert(str(temp_pdf_path))
            _docling_log("Page conversion completed")
            doc = conv_result.document

            _docling_log("Exporting page to markdown...")
            markdown = doc.export_to_markdown(image_mode=ImageRefMode.REFERENCED)

            if markdown:
                char_count = len(markdown.strip())
                _docling_log(f"Extraction successful: {char_count} characters extracted")
                _docling_log("Markdown content:")
                _docling_log("=" * 80)
                lines = markdown.strip().split('\n')
                for i, line in enumerate(lines[:50], 1):  # Show first 50 lines
                    _docling_log(line)
                if len(lines) > 50:
                    _docling_log(f"... ({len(lines) - 50} more lines)")
                _docling_log("=" * 80)

                # Extract instrument IDs from markdown
                import re
                instrument_ids = re.findall(r'\b[A-Z]{1,3}-?\d{1,4}\b', markdown)
                if instrument_ids:
                    _docling_log(f"Detected instrument IDs in OCR: {', '.join(set(instrument_ids[:20]))}")
                    if len(set(instrument_ids)) > 20:
                        _docling_log(f"  ... and {len(set(instrument_ids)) - 20} more")
                else:
                    _docling_log("⚠️  No instrument IDs detected in OCR output")

                _docling_log(f"Markdown context ready: {char_count} chars will be sent to AI model")
                return markdown.strip()
            else:
                _docling_log("Extraction completed but no markdown was generated")
                return None
        finally:
            if temp_pdf_path.exists():
                temp_pdf_path.unlink()
    except Exception as exc:
        _docling_log(f"ERROR: {exc}")
        import traceback
        _docling_log(f"Traceback: {traceback.format_exc()}")
        return None
    finally:
        for logger, h in handlers:
            logger.removeHandler(h)
            h.close()


def write_instruments_to_excel(transcription_json: str, excel_path: Path) -> None:
    """Write the instruments list from a transcription JSON to an Excel workbook."""
    if openpyxl is None:
        raise RuntimeError("openpyxl is required for Excel export. Install it with: pip install openpyxl")
    data = json.loads(transcription_json)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Instruments"
    headers = ["Page", "Name", "Description", "Responsibility", "Wire Count", "Wire Type", "Voltage", "Intrinsic Safety", "Notes"]
    ws.append(headers)
    for page_result in data.get("pages", []):
        page_num = page_result.get("page", "")
        for instrument in page_result.get("instruments", []):
            ws.append([
                page_num,
                instrument.get("name", ""),
                instrument.get("description", ""),
                instrument.get("Responsibility", ""),
                instrument.get("wire_count", ""),
                instrument.get("wire_type", ""),
                instrument.get("voltage", ""),
                instrument.get("intrinsic_safety", ""),
                instrument.get("Controls Notes") or instrument.get("P&ID Notes") or instrument.get("Notes", ""),
            ])
    # Auto-size columns
    for col in ws.columns:
        max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
    wb.save(str(excel_path))


COMBINED_EXCEL_COLUMNS = [
    ("name", "Name"), ("description", "Description"), ("Responsibility", "Responsibility"),
    ("wire_count", "Wire Count"), ("wire_type", "Wire Type"), ("voltage", "Voltage"),
    ("intrinsic_safety", "Intrinsic Safety"), ("source", "Source"),
    ("controls_responsibility", "Controls Responsibility"), ("pids_responsibility", "P&ID Responsibility"),
    ("conflict_notes", "Conflict Notes"), ("controls_notes", "Controls Notes"), ("pids_notes", "P&ID Notes"),
    ("controls_pages", "Controls Pages"), ("pids_pages", "P&ID Pages"),
]


def write_combined_to_excel(combined: dict, excel_path: Path) -> None:
    """Write the combined instrument list to Excel; rows with a responsibility conflict are filled red."""
    if openpyxl is None:
        raise RuntimeError("openpyxl is required for Excel export. Install it with: pip install openpyxl")
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Combined"
    ws.append([header for _, header in COMBINED_EXCEL_COLUMNS])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    red = PatternFill("solid", fgColor="FFC7CE")
    for row in combined.get("combined_instruments", []):
        if not isinstance(row, dict):
            continue
        values = [", ".join(map(str, v)) if isinstance(v, list) else v for v in (row.get(key, "") for key, _ in COMBINED_EXCEL_COLUMNS)]
        ws.append(values)
        if row.get("responsibility_conflict"):
            for cell in ws[ws.max_row]:
                cell.fill = red
    for col in ws.columns:
        max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
    ws.freeze_panes = "B2"
    wb.save(str(excel_path))


def _normalize_host(api_host: str, default: str) -> str:
    """Normalize a host so both :11434 and :11434/v1 are supported for Ollama-style servers."""
    normalized = (api_host or default).rstrip("/")
    if normalized.endswith("/v1"):
        normalized = normalized[:-3]
    return normalized


def _is_claude_model(model_name: Optional[str]) -> bool:
    if not model_name or is_claude_code_model(model_name):
        return False  # "Claude Code: ..." models go through the CLI, not the Anthropic API
    lowered = model_name.lower()
    return "claude" in lowered or "anthropic" in lowered


CLOUD_PROVIDER_CHOICES = ["Local", "Anthropic", "OpenAI", "Google", "Custom"]


def _normalize_cloud_provider(provider: Optional[str]) -> str:
    if not provider:
        return "Local"
    lowered = (provider or "").strip().lower()
    if lowered in {"local", "ollama"}:
        return "Local"
    if lowered in {"anthropic", "claude", "claudeai"}:
        return "Anthropic"
    if lowered in {"openai", "gpt"}:
        return "OpenAI"
    if lowered in {"google", "gemini", "vertex"}:
        return "Google"
    return "Custom"


def render_pdf_pages_to_images(pdf_path: Path, dpi: int = 300, output_dir: Optional[Path] = None) -> List[Path]:
    """Render each PDF page to a full-resolution PNG file."""
    if fitz is None:
        raise RuntimeError(
            "PyMuPDF is required. Install it with: pip install pymupdf"
        ) from _FITZ_IMPORT_ERROR

    doc = fitz.open(str(pdf_path))
    if output_dir is None:
        output_dir = Path(tempfile.mkdtemp(prefix="pdf_pages_", dir=str(pdf_path.parent)))
    else:
        output_dir = output_dir.expanduser().resolve()

    output_dir.mkdir(parents=True, exist_ok=True)
    image_paths: List[Path] = []
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    for page_number, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        image_path = output_dir / f"page_{page_number:03d}.png"
        pix.save(str(image_path))
        image_paths.append(image_path)

    return image_paths


def _load_image_bytes(image_path: Path) -> str:
    with image_path.open("rb") as handle:
        return base64.b64encode(handle.read()).decode("utf-8")


def _read_ollama_stream(response: "requests.Response", on_chunk: Callable[[str], None]) -> dict:
    """Consume an Ollama /api/chat NDJSON stream, forwarding text to on_chunk; return a non-streamed-shaped result."""
    content: List[str] = []
    thinking: List[str] = []
    try:
        # chunk_size=1: requests otherwise buffers 512 bytes before yielding a line, which defeats streaming.
        for raw in response.iter_lines(chunk_size=1):
            if not raw:
                continue
            chunk = json.loads(raw)
            if chunk.get("error"):
                raise RuntimeError(f"Ollama error: {chunk['error']}")
            message = chunk.get("message", {})
            if message.get("thinking"):
                thinking.append(message["thinking"])
                on_chunk(message["thinking"])
            if message.get("content"):
                content.append(message["content"])
                on_chunk(message["content"])
            if chunk.get("done"):
                break
    finally:
        response.close()
    on_chunk("\n")
    message = {"content": "".join(content)}
    if thinking:
        message["thinking"] = "".join(thinking)
    return {"message": message, "done": True}


def _read_claude_stream(response: "requests.Response", on_chunk: Callable[[str], None]) -> dict:
    """Consume an Anthropic Messages server-sent-event stream, forwarding text to on_chunk."""
    text: List[str] = []
    thinking: List[str] = []
    try:
        for raw_line in response.iter_lines(chunk_size=1):
            raw = raw_line.decode("utf-8") if isinstance(raw_line, bytes) else raw_line
            if not raw.startswith("data:"):
                continue
            event = json.loads(raw[5:].strip())
            if event.get("type") == "error":
                raise RuntimeError(f"Claude stream error: {event.get('error')}")
            if event.get("type") != "content_block_delta":
                continue
            delta = event.get("delta", {})
            if delta.get("type") == "text_delta":
                text.append(delta["text"])
                on_chunk(delta["text"])
            elif delta.get("type") == "thinking_delta":
                thinking.append(delta["thinking"])
                on_chunk(delta["thinking"])
    finally:
        response.close()
    on_chunk("\n")
    blocks: List[dict] = []
    if thinking:
        blocks.append({"type": "thinking", "thinking": "".join(thinking)})
    blocks.append({"type": "text", "text": "".join(text)})
    return {"content": blocks}


def _send_ollama_request(
    model: str,
    api_host: str,
    api_key: str,
    prompt: str,
    image_paths: List[Path],
    max_tokens: int,
    log_callback: Optional[Callable[[str], None]] = None,
    docling_markdown: Optional[str] = None,
    stream_callback: Optional[Callable[[str], None]] = None,
) -> str:
    def _log(message: str) -> None:
        if log_callback is not None:
            log_callback(message)

    # Every model request streams; with no listener the chunks are simply discarded.
    stream_callback = stream_callback or (lambda _text: None)

    headers = {
        "Content-Type": "application/json",
    }
    if api_key and api_key != "None":
        headers["Authorization"] = f"Bearer {api_key}"

    encoded_images = [_load_image_bytes(path) for path in image_paths]

    prompt_with_context = prompt
    if docling_markdown:
        prompt_with_context = f"{prompt}\n\n---\nStructured document context (Docling extraction):\n{docling_markdown}"

    payload = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": "You are a careful engineering document analyst. Return strict JSON for the requested extraction.",
            },
            {
                "role": "user",
                "content": prompt_with_context,
                "images": encoded_images,
            },
        ],
        "stream": True,
        "options": {
            "num_predict": max_tokens,
            "num_ctx": int(os.getenv("OLLAMA_NUM_CTX", max_tokens)),
        },
    }

    # Enable thinking/verbose output for models that support it
    if os.getenv("OLLAMA_ENABLE_THINKING", "").lower() in {"true", "1", "yes"}:
        payload["thinking"] = "enabled"
        _log("[INFO] Extended thinking enabled for this request")

    try:
        response = requests.post(
            f"{_normalize_host(api_host, 'http://localhost:11434')}/api/chat",
            headers=headers,
            data=json.dumps(payload),
            stream=True,
            # timeout=300,
        )
        response.raise_for_status()
        result = _read_ollama_stream(response, stream_callback)
    except requests.exceptions.ConnectionError as exc:
        raise RuntimeError(
            f"Could not reach Ollama at {api_host or 'http://localhost:11434'}. Start the Ollama service and ensure it is listening on /api/chat."
        ) from exc
    except requests.exceptions.Timeout as exc:
        raise RuntimeError(f"Timed out while contacting Ollama at {api_host or 'http://localhost:11434'}.") from exc
    except requests.HTTPError as exc:
        error_body = None
        try:
            error_body = exc.response.text if exc.response is not None else None
        except Exception:
            error_body = None

        details = f"Ollama request failed: {exc}"
        if error_body:
            details += f". Response body: {error_body}"
            lowered = error_body.lower()
            if "image" in lowered or "vision" in lowered:
                details += " | Hint: the selected model may not support image inputs; choose a vision-capable model."
            if "unknown field" in lowered:
                details += " | Hint: your Ollama server rejected one or more request fields; update Ollama or adjust request format."
        raise RuntimeError(details) from exc

    message = result.get("message", {})
    content = message.get("content", "")

    _log(f"Response received with keys: {', '.join(result.keys())}")
    _log(f"Message contains: {', '.join(message.keys())}")

    thinking = None
    if "thinking" in message:
        thinking = message.get("thinking", "")
        _log("Extracting model thinking process...")
    elif "reasoning" in message:
        thinking = message.get("reasoning", "")
        _log("Extracting model reasoning...")
    elif "thought" in message:
        thinking = message.get("thought", "")
        _log("Extracting model thoughts...")

    if not thinking and "thinking" in result:
        thinking = result.get("thinking", "")
        _log("Found thinking in response...")

    if thinking and isinstance(thinking, str) and thinking.strip():
        _log(f"[MODEL THINKING]\n{thinking}")
    else:
        _log(" No thinking field found in response")

    # Extract thinking from content if it has XML-style markers
    if not thinking and content and isinstance(content, str):
        import re
        thinking_match = re.search(r'<thinking>(.*?)</thinking>', content, re.DOTALL)
        if thinking_match:
            thinking = thinking_match.group(1).strip()
            _log(f"[MODEL THINKING - from content]\n{thinking}")

    if not isinstance(content, str):
        raise RuntimeError(f"Unexpected Ollama response format: {result}")
    return content.strip()


def _send_claude_request(
    model: str,
    api_host: str,
    api_key: str,
    prompt: str,
    image_paths: List[Path],
    max_tokens: int,
    log_callback: Optional[Callable[[str], None]] = None,
    docling_markdown: Optional[str] = None,
    stream_callback: Optional[Callable[[str], None]] = None,
) -> str:
    if not api_key or api_key == "None":
        raise RuntimeError("A Claude API key is required. Set AI_GATEWAY_API_KEY or pass --api-key.")

    def _log(message: str) -> None:
        if log_callback is not None:
            log_callback(message)

    # Every model request streams; with no listener the chunks are simply discarded.
    stream_callback = stream_callback or (lambda _text: None)

    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }

    prompt_with_context = prompt
    if docling_markdown:
        prompt_with_context = f"{prompt}\n\n---\nStructured document context (Docling extraction):\n{docling_markdown}"

    content_blocks = [{"type": "text", "text": prompt_with_context}]
    for image_path in image_paths:
        encoded_image = _load_image_bytes(image_path)
        content_blocks.append(
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": encoded_image,
                },
            }
        )

    safe_max_tokens = max(1, max_tokens)
    available_models = _get_claude_models(api_key, api_host)
    if available_models and model not in available_models:
        available_preview = ", ".join(available_models[:8])
        raise RuntimeError(
            f"Claude model '{model}' is not available from {api_host}. Choose one of the discovered models: {available_preview}"
        )

    payload = {
        "model": model,
        "max_tokens": safe_max_tokens,
        "system": "You are an engineering document analyst. Extract the PLC-connected instruments and return valid JSON.",
        "messages": [
            {
                "role": "user",
                "content": content_blocks,
            }
        ],
    }

    if os.getenv("CLAUDE_ENABLE_THINKING", "").lower() in {"true", "1", "yes"}:
        payload["thinking"] = {"type": "enabled", "budget_tokens": min(10000, safe_max_tokens // 2)}
        _log("[INFO] Extended thinking enabled for Claude")

    payload["stream"] = True

    base_host = (api_host or "https://api.anthropic.com").rstrip("/")
    if base_host.endswith("/v1"):
        url = f"{base_host}/messages"
    else:
        url = f"{base_host}/v1/messages"

    try:
        response = requests.post(url, headers=headers, data=json.dumps(payload), stream=True)
                                #  timeout=300)
        response.raise_for_status()
        result = _read_claude_stream(response, stream_callback)
    except requests.HTTPError as exc:
        error_body = None
        try:
            error_body = response.text  # type: ignore[name-defined]
        except Exception:
            error_body = None
        raise RuntimeError(f"Claude request failed for model '{model}': {exc}. Response body: {error_body}") from exc

    content = result.get("content", [])
    if isinstance(content, list):
        text_parts = []
        thinking_text = []
        for item in content:
            if isinstance(item, dict):
                if item.get("type") == "thinking":
                    thinking = item.get("thinking", "").strip()
                    if thinking:
                        thinking_text.append(thinking)
                elif item.get("type") == "text":
                    text_parts.append(item.get("text", ""))

        if thinking_text:
            _log(f"[CLAUDE THINKING]\n{''.join(thinking_text)}")

        return "\n".join(text_parts).strip()
    raise RuntimeError(f"Unexpected Claude response format: {result}")


def _extract_complete_objects(text: str) -> List[dict]:
    """Extract all complete top-level JSON objects from a string, tolerating a truncated tail.
    
    Scans past any outer unclosed wrapper to find individually-complete objects.
    """
    objects: List[dict] = []
    i = 0
    while i < len(text):
        # Skip to next opening brace
        start = text.find("{", i)
        if start == -1:
            break
        depth = 0
        in_string = False
        escape = False
        end = -1
        for j in range(start, len(text)):
            ch = text[j]
            if escape:
                escape = False
                continue
            if ch == "\\" and in_string:
                escape = True
                continue
            if ch == '"':
                in_string = not in_string
                continue
            if in_string:
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    end = j
                    break
        if end == -1:
            # This object was never closed — truncated. Try to find objects at shallower nesting.
            # Look for inner objects that were completed (depth goes to 0 within the range)
            i = start + 1
            continue
        candidate = text[start : end + 1]
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                objects.append(parsed)
            i = end + 1
        except json.JSONDecodeError:
            i = start + 1
    return objects


def _extract_json_payload(response_text: str) -> dict:
    """Extract a JSON object from a model response that may include fenced code blocks or surrounding prose.
    
    If the response is truncated mid-JSON, salvages any complete instrument objects found.
    """
    if not response_text:
        return {"plc": "[unclear]", "instruments": []}

    cleaned = response_text.strip()
    if cleaned.startswith("```"):
        lines = [line.strip() for line in cleaned.splitlines() if line.strip()]
        if len(lines) >= 2 and lines[0].startswith("```"):
            fenced = []
            for line in lines[1:]:
                if line.startswith("```"):
                    break
                fenced.append(line)
            cleaned = "\n".join(fenced)

    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidate = cleaned[start : end + 1]
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            pass

    try:
        parsed = json.loads(cleaned)
        if isinstance(parsed, dict):
            return parsed
    except json.JSONDecodeError:
        pass

    # JSON is malformed (likely truncated) — salvage complete instrument objects
    all_objects = _extract_complete_objects(cleaned)
    # The outermost wrapper object may be incomplete; collect instrument-like dicts
    instruments: List[dict] = []
    plc = "[unclear]"
    for obj in all_objects:
        if "instruments" in obj:
            # This is a wrapper object that was parsed fully
            return obj
        if "name" in obj or "wire_count" in obj:
            # Looks like an individual instrument entry
            if "plc" in obj:
                plc = obj["plc"]
            else:
                instruments.append(obj)
        elif "plc" in obj and isinstance(obj.get("plc"), str):
            plc = obj["plc"]
    if instruments:
        return {"plc": plc, "instruments": instruments, "Notes": "Response was truncated; partial results recovered."}
    return {"plc": "[unclear]", "instruments": []}


def _build_page_payload(page_number: int, page_result: dict) -> dict:
    """Wrap a single-page AI response in a page-specific section for the final JSON output."""
    payload = {"page": page_number}
    payload.update(page_result)
    return payload


def _deduplicate_instruments(instruments: List[dict]) -> List[dict]:
    """Remove repeated instrument entries while preserving the first occurrence."""
    seen: set = set()
    deduped: List[dict] = []
    for instrument in instruments:
        if not isinstance(instrument, dict):
            continue
        key = json.dumps(instrument, sort_keys=True, ensure_ascii=False)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(instrument)
    return deduped


def _assemble_transcription_output(page_results: List[dict]) -> dict:
    """Combine per-page findings into a single JSON document with a pages section."""
    combined: dict = {"pages": page_results}
    plc_name = "[unclear]"
    for page_result in page_results:
        page_plc = page_result.get("plc")
        if isinstance(page_plc, str) and page_plc and page_plc != "[unclear]":
            plc_name = page_plc
    combined["plc"] = plc_name
    return combined


REFERENCE_IMAGE_NOTE = (
    "\n\nIMPORTANT - REFERENCE IMAGE: The FIRST image attached is only a REFERENCE: a legend showing how symbols "
    "and line styles on this kind of drawing are drawn and what they mean. It is NOT part of the drawing to analyze. "
    "Use it only to understand notation. Never list, count or transcribe any tag, symbol, device or text from it, "
    "and do not mention it in the output. The LAST image attached is the actual drawing page; every instrument, "
    "connection and note you report must come from that page."
)


def with_reference_note(prompt: str, has_reference: bool) -> str:
    """Append the reference-image explanation when a reference image is attached ahead of the page image."""
    return prompt + REFERENCE_IMAGE_NOTE if has_reference else prompt


CARRY_FORWARD_INSTRUCTIONS = (
    "\n\nCARRY-FORWARD NOTE: pages of this drawing set are analyzed one at a time, and you will not see this page again. "
    "In addition to the instruments, add a top-level key \"carry_forward\" to your JSON: a list of short strings (one fact each) "
    "holding critical information from THIS page that a later page will need but cannot see for itself, such as:\n"
    "- voltages, power sources or rails on terminals/terminal strips and which devices they feed\n"
    "- descriptions or labels of circuits, cables, junction boxes or signals that continue to or arrive from other pages "
    "(include the off-page reference, e.g. 'wire 709B continues to sheet 14')\n"
    "- Anything that starts on this page and leaves to continue on other pages.\n"
    "- the PLC/panel name, module or slot assignments, and drawing-wide notes, legends or conventions that apply to other pages\n"
)


def with_carry_forward(prompt: str, notes: List[str]) -> str:
    """Append the carry-forward instructions, and the notes gathered from earlier pages when there are any."""
    text = prompt + CARRY_FORWARD_INSTRUCTIONS
    if notes:
        text += ("\n\nNOTES CARRIED FORWARD FROM EARLIER PAGES (information that is not visible on this page; use it to "
                 "interpret this page, but only report instruments that appear on this page):\n"
                 + "\n".join(f"- {note}" for note in notes))
    return text


def update_carry_forward(notes: List[str], page_number: int, page_result: dict) -> int:
    """Append the new items a page returned under 'carry_forward' to the running notes; returns how many were added."""
    items = page_result.get("carry_forward")
    if isinstance(items, str):
        items = [items]
    added = 0
    known = {re.sub(r"^p\d+: ", "", n).strip().lower() for n in notes}
    for item in items if isinstance(items, list) else []:
        text = str(item).strip()
        if not text or text.lower() in known:
            continue
        known.add(text.lower())
        notes.append(f"p{page_number}: {text}")
        added += 1
    return added


def _render_controls_page_prompt(page_number: int, total_pages: int) -> str:
        return f"""You are analyzing controls wiring schematics.

Focus only on instruments/devices connected to the PLC from external locations and all PLC-related internal/external tags shown on this controls page.
On wiring schematics, external customer-wired connections are dashed lines. Dashed-dot lines indicate SKID responsibility. 
Use the reference image provided as a guide to the drawings design standards.
Use BOTH the visual diagram AND the extracted text to ensure complete identification.

Return strict JSON in this structure:
{{
    "instruments": [
        {{
            "name": "instrument tag",
            "description": "instrument description if available, otherwise [unclear]",
            "Responsibility": "CUSTOMER, SKID, PLC, or [unclear]",
            "wire_count": "integer count of the wires shown on the drawing or [unclear]",
            "wire_type": "shielded/unshielded/coax/fiber/power/other/[unclear]",
            "voltage": "bare value only (e.g. 24VDC, 120VAC) and only if labelled on the drawing for this device's own circuit; do NOT infer it from a power rail on the sheet or from the device type, use [unclear] and explain any assumption in the notes",
            "intrinsic_safety": true or false,
            "Controls Notes": "ambiguities or assumptions"
        }}
    ]
}}
"""


def _render_pids_page_prompt(page_number: int, total_pages: int) -> str:
        return f"""You are analyzing a P&ID drawing.

IMPORTANT: Refer to the extracted text context (provided below the image) to identify all instrument IDs and tags.
Match the visual circles/tags in the diagram with the instrument IDs found in the extracted text.
Use BOTH the visual diagram AND the extracted text to ensure complete identification.

For this P&ID step, do not infer wire count, wire type, or intrinsic safety unless explicitly shown.

Return strict JSON in this structure:
{{
    "instruments": [
        {{
            "name": "instrument tag",
            "Responsibility": "CUSTOMER, SKID, PLC, or [unclear]",
            "P&ID Notes": "ambiguities or assumptions"
        }}
    ]
}}

If something is unclear, use [unclear]. Do not invent values.
"""


def parse_page_selection(selection: str, total_pages: int) -> List[int]:
    """Parse a page selection string into a sorted list of 1-based page numbers.
    
    Accepts: single page ("3"), ranges ("2-5"), comma-separated ("1,3,5"),
    and combinations ("1,3-5,7"). Returns all pages if selection is empty.
    """
    if not selection or not selection.strip():
        return list(range(1, total_pages + 1))
    pages: set[int] = set()
    for part in selection.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            bounds = part.split("-", 1)
            start, end = int(bounds[0].strip()), int(bounds[1].strip())
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))
    return sorted(p for p in pages if 1 <= p <= total_pages)


def transcribe_pdf_stage(
    pdf_path: Path,
    prompt_builder: Callable[[int, int], str],
    model: Optional[str] = None,
    api_host: Optional[str] = None,
    api_key: Optional[str] = None,
    max_tokens: int = 2048,
    dpi: int = 300,
    pages: Optional[List[int]] = None,
    reference_image: Optional[Path] = None,
    log_callback: Optional[callable] = None,
    stream_callback: Optional[Callable[[str], None]] = None,
    stage_name: str = "stage",
    docling_enable: bool = True,
    docling_do_ocr: bool = True,
    docling_do_table_structure: bool = True,
    docling_do_picture_description: bool = False,
    docling_images_scale: float = 1.0,
) -> str:
    """Render each PDF page to full-resolution PNG images and send them to an AI model one page at a time."""
    image_paths = render_pdf_pages_to_images(pdf_path, dpi=dpi)
    total_pages = len(image_paths)
    selected_pages = pages if pages else list(range(1, total_pages + 1))

    resolved_model = model or os.getenv("AI_GATEWAY_MODEL", "gemma4:e2b")
    resolved_host = api_host or os.getenv("AI_GATEWAY_HOST", "http://localhost:11434")
    resolved_api_key = api_key or os.getenv("AI_GATEWAY_API_KEY") or "None"

    def _log(message: str) -> None:
        if log_callback is not None:
            log_callback(message)

    page_results: List[dict] = []
    carry_notes: List[str] = []  # running note: each page's model adds what later pages will need but cannot see
    _log(f"Starting {stage_name} using model '{resolved_model}'")
    _log(f"Docling: enabled={docling_enable}, OCR={docling_do_ocr}, Tables={docling_do_table_structure}, Pictures={docling_do_picture_description}, Scale={docling_images_scale}x")

    if docling_enable:
        _log(f"OCR pre-step: Docling full-page OCR (RapidOCR) at {docling_images_scale}x, text sent to the model as context")
        if docling_images_scale < 3.0:
            _log(f"TIP: small instrument tags need a high OCR scale ({docling_images_scale}x now); 3.0-4.0 works well")
    else:
        _log("OCR pre-step: off (the model only gets the page image)")
    if reference_image and reference_image.is_file():
        _log(f"Reference image (notation legend only, sent first with each page): {reference_image}")
    else:
        _log("Reference image: none")

    # Pre-step: full-page OCR of every selected page, done once up front (cached on disk).
    docling_markdown_cache: dict[int, Optional[str]] = {}
    if docling_enable:
        docling_markdown_cache.update(preprocess_pdf_ocr(
            pdf_path,
            pages=selected_pages,
            do_table_structure=docling_do_table_structure,
            images_scale=docling_images_scale,
            log_callback=_log,
        ))

    for index, image_path in enumerate(image_paths, start=1):
        if index not in selected_pages:
            continue
        _log(f"Processing page {index}/{total_pages}")
        has_reference = bool(reference_image and reference_image.is_file())
        prompt = with_carry_forward(with_reference_note(prompt_builder(index, len(image_paths)), has_reference), carry_notes)
        if carry_notes:
            _log(f"Carrying {len(carry_notes)} note(s) from earlier pages into page {index}:\n" + "\n".join(f"  - {n}" for n in carry_notes))
        page_images = ([reference_image] if has_reference else []) + [image_path]

        docling_markdown = docling_markdown_cache.get(index)
        if docling_enable and not docling_markdown:
            _log(f"[OCR PRE-STEP] No OCR text for page {index}; continuing with the image only")

        response_text = ""
        for attempt in range(1, 4):
            try:
                backend = ("Claude Code CLI" if is_claude_code_model(resolved_model)
                           else "Anthropic API" if _is_claude_model(resolved_model) else "Ollama/OpenAI-compatible API")
                image_note = ", ".join(p.name for p in page_images)
                if has_reference:
                    image_note += " (first = reference legend, last = the page)"
                ocr_note = f"{len(docling_markdown)} chars of OCR text" if docling_markdown else "no OCR text"
                _log(f"Sending page {index} to {resolved_model} via {backend}: {len(page_images)} image(s) [{image_note}], {ocr_note}")
                _log("Model processing started...")

                if is_claude_code_model(resolved_model):
                    response_text = _claude_code_send(
                        resolved_model, prompt, page_images,
                        docling_markdown=docling_markdown, log_callback=_log, stream_callback=stream_callback,
                    )
                elif _is_claude_model(resolved_model):
                    response_text = _send_claude_request(
                        model=resolved_model,
                        api_host=resolved_host,
                        api_key=resolved_api_key,
                        prompt=prompt,
                        image_paths=page_images,
                        max_tokens=max_tokens,
                        log_callback=_log,
                        docling_markdown=docling_markdown,
                        stream_callback=stream_callback,
                    )
                else:
                    response_text = _send_ollama_request(
                        model=resolved_model,
                        api_host=resolved_host,
                        api_key=resolved_api_key,
                        prompt=prompt,
                        image_paths=page_images,
                        max_tokens=max_tokens,
                        log_callback=_log,
                        docling_markdown=docling_markdown,
                        stream_callback=stream_callback,
                    )
            except Exception as exc:
                _log(f"Page {index} attempt {attempt} failed with error: {exc}")
                response_text = ""
                if attempt < 3:
                    is_overload = "529" in str(exc) or "overload" in str(exc).lower()
                    wait_seconds = (30 * attempt) if is_overload else (5 * attempt)
                    _log(f"Waiting {wait_seconds}s before retry...")
                    import time as _time
                    _time.sleep(wait_seconds)

            if response_text.strip():
                _log("Model processing completed.")
                _log(f"Response received: {len(response_text)} characters")
                if attempt > 1:
                    _log(f"Page {index} succeeded on attempt {attempt}.")
                break
            _log(f"Page {index} attempt {attempt} returned a blank response{', retrying...' if attempt < 3 else ', giving up.'}")

        _log(f"Raw model response for page {index}: {response_text}")
        parsed_result = _extract_json_payload(response_text)
        added = update_carry_forward(carry_notes, index, parsed_result)
        _log(f"Page {index} added {added} carry-forward note(s); running note is {len(carry_notes)} item(s)")
        page_results.append(_build_page_payload(index, parsed_result))
        _log(f"Parsed page {index}: {json.dumps(parsed_result, ensure_ascii=False)}")

    combined_output = _assemble_transcription_output(page_results)
    _log(f"Final combined output: {json.dumps(combined_output, ensure_ascii=False)}")
    return json.dumps(combined_output, indent=2, ensure_ascii=False)


def transcribe_controls_drawings(
    pdf_path: Path,
    model: Optional[str] = None,
    api_host: Optional[str] = None,
    api_key: Optional[str] = None,
    max_tokens: int = 2048,
    dpi: int = 300,
    pages: Optional[List[int]] = None,
    reference_image: Optional[Path] = None,
    log_callback: Optional[callable] = None,
    stream_callback: Optional[Callable[[str], None]] = None,
    docling_enable: bool = True,
    docling_do_ocr: bool = True,
    docling_do_table_structure: bool = True,
    docling_do_picture_description: bool = False,
    docling_images_scale: float = 1.0,
) -> str:
    return transcribe_pdf_stage(
        pdf_path=pdf_path,
        prompt_builder=_render_controls_page_prompt,
        model=model,
        api_host=api_host,
        api_key=api_key,
        max_tokens=max_tokens,
        dpi=dpi,
        pages=pages,
        reference_image=reference_image,
        log_callback=log_callback,
        stream_callback=stream_callback,
        stage_name="controls extraction",
        docling_enable=docling_enable,
        docling_do_ocr=docling_do_ocr,
        docling_do_table_structure=docling_do_table_structure,
        docling_do_picture_description=docling_do_picture_description,
        docling_images_scale=docling_images_scale,
    )


def transcribe_pids_drawings(
    pdf_path: Path,
    model: Optional[str] = None,
    api_host: Optional[str] = None,
    api_key: Optional[str] = None,
    max_tokens: int = 2048,
    dpi: int = 300,
    pages: Optional[List[int]] = None,
    reference_image: Optional[Path] = None,
    log_callback: Optional[callable] = None,
    stream_callback: Optional[Callable[[str], None]] = None,
    docling_enable: bool = True,
    docling_do_ocr: bool = True,
    docling_do_table_structure: bool = True,
    docling_do_picture_description: bool = False,
    docling_images_scale: float = 1.0,
) -> str:
    return transcribe_pdf_stage(
        pdf_path=pdf_path,
        prompt_builder=_render_pids_page_prompt,
        model=model,
        api_host=api_host,
        api_key=api_key,
        max_tokens=max_tokens,
        dpi=dpi,
        pages=pages,
        reference_image=reference_image,
        log_callback=log_callback,
        stream_callback=stream_callback,
        stage_name="p&id extraction",
        docling_enable=docling_enable,
        docling_do_ocr=docling_do_ocr,
        docling_do_table_structure=docling_do_table_structure,
        docling_do_picture_description=docling_do_picture_description,
        docling_images_scale=docling_images_scale,
    )


def _extract_json_object(response_text: str) -> dict:
    """Best-effort extraction of a single JSON object from model output."""
    if not response_text:
        return {}

    cleaned = response_text.strip()
    if cleaned.startswith("```"):
        lines = [line.rstrip() for line in cleaned.splitlines()]
        stripped: List[str] = []
        in_fence = False
        for line in lines:
            if line.strip().startswith("```"):
                if in_fence:
                    break
                in_fence = True
                continue
            if in_fence:
                stripped.append(line)
        if stripped:
            cleaned = "\n".join(stripped).strip()

    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidate = cleaned[start : end + 1]
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            pass

    try:
        parsed = json.loads(cleaned)
        if isinstance(parsed, dict):
            return parsed
    except json.JSONDecodeError:
        pass

    objects = _extract_complete_objects(cleaned)
    if objects:
        return objects[-1]
    return {}


_TAG_TOKEN = re.compile(r"[A-Z]{1,5}[-_ ]?\d+[A-Z]?(?![A-Za-z0-9])")
_NON_INSTRUMENT = re.compile(r"^(PLC\b|FROM\b|REMOTE\b|SYSTEM\b|(AI|AO|DI|DO|RTD|CNTR)\b)|\b(POWER|FEED|SUPPLY)\b", re.IGNORECASE)


def _is_unclear(value: object) -> bool:
    return value is None or str(value).strip().lower() in {"", "[unclear]", "unclear", "none", "null"}


def _tag_keys(name: str) -> List[str]:
    """Normalized join keys for a tag: every tag-like token in the name, so 'SV1210 (XV1210)' yields SV1210 and XV1210."""
    upper = str(name).upper()
    keys = [re.sub(r"[^A-Z0-9]", "", match.group(0)) for match in _TAG_TOKEN.finditer(upper)]
    return list(dict.fromkeys(keys)) or [re.sub(r"[^A-Z0-9]", "", upper)]


def _is_non_instrument(name: str) -> Optional[str]:
    """Reason a row is not a field instrument (module placeholder, power feed, ...), else None."""
    if _NON_INSTRUMENT.search(str(name)):
        return "power feed / PLC module / system signal row"
    if not re.search(r"\d", str(name)):
        return "no tag number"
    return None


def _flatten_stage(payload: dict) -> List[dict]:
    """All instrument dicts from a stage output, tagged with their page number."""
    rows: List[dict] = []
    pages = payload.get("pages")
    if isinstance(pages, list):
        for page in pages:
            for inst in page.get("instruments", []) if isinstance(page, dict) else []:
                if isinstance(inst, dict) and inst.get("name"):
                    rows.append({**inst, "_page": page.get("page")})
    for inst in payload.get("instruments", []) if isinstance(payload.get("instruments"), list) else []:
        if isinstance(inst, dict) and inst.get("name"):
            rows.append({**inst, "_page": None})
    return rows


def _collapse_by_key(rows: List[dict]) -> dict:
    """One record per primary tag key; repeats (the same tag on several pages) merge, keeping known values."""
    collapsed: dict = {}
    for row in rows:
        key = _tag_keys(row["name"])[0]
        record = collapsed.get(key)
        if record is None:
            collapsed[key] = {**row, "_pages": [row["_page"]] if row["_page"] is not None else [],
                              "_keys": _tag_keys(row["name"])}
            continue
        for field, value in row.items():
            if field.startswith("_"):
                continue
            if _is_unclear(record.get(field)) and not _is_unclear(value):
                record[field] = value
        if row["_page"] is not None and row["_page"] not in record["_pages"]:
            record["_pages"].append(row["_page"])
    return collapsed


def _ask_model_to_pair(
    candidates: List[tuple],
    model: str,
    api_host: str,
    api_key: str,
    max_tokens: int,
    log_callback: Optional[Callable[[str], None]],
    stream_callback: Optional[Callable[[str], None]],
) -> List[tuple]:
    """Ask the model which near-miss (controls tag, P&ID tag) pairs are the same device; only these short pairs are sent."""
    def _log(message: str) -> None:
        if log_callback is not None:
            log_callback(message)

    listing = "\n".join(
        f'{i}. controls "{c["name"]}" ({c.get("description", "")}) <-> P&ID "{p["name"]}" ({p.get("description", "")})'
        for i, (c, p) in enumerate(candidates)
    )
    prompt = (
        "Below are candidate pairs of instrument tags, one from a controls wiring drawing and one from a P&ID.\n"
        "Decide for each pair whether both tags name the SAME physical device (tags often differ only by punctuation, "
        "a suffix letter, or a typo). Be conservative: answer false when unsure.\n"
        'Reply with strict JSON only: {"same": [<index>, ...]} listing the indexes that are the same device.\n\n'
        f"{listing}"
    )
    response_text = ""
    for attempt in range(1, 4):
        try:
            if is_claude_code_model(model):
                response_text = _claude_code_send(model, prompt, [], log_callback=_log, stream_callback=stream_callback)
            elif _is_claude_model(model):
                response_text = _send_claude_request(model=model, api_host=api_host, api_key=api_key, prompt=prompt,
                                                     image_paths=[], max_tokens=max_tokens, log_callback=_log,
                                                     docling_markdown=None, stream_callback=stream_callback)
            else:
                response_text = _send_ollama_request(model=model, api_host=api_host, api_key=api_key, prompt=prompt,
                                                     image_paths=[], max_tokens=max_tokens, log_callback=_log,
                                                     docling_markdown=None, stream_callback=stream_callback)
        except Exception as exc:
            _log(f"Pairing attempt {attempt} failed: {exc}")
            response_text = ""
        if response_text.strip():
            break
        _log(f"Pairing attempt {attempt} returned blank output{', retrying...' if attempt < 3 else ', giving up.'}")
    same = _extract_json_object(response_text).get("same", [])
    return [candidates[i] for i in same if isinstance(i, int) and 0 <= i < len(candidates)]


def _merge_pair(controls: Optional[dict], pids: Optional[dict]) -> dict:
    """Merge one device's records. Everything except tag and Responsibility comes from controls (the P&ID has no
    wiring data); Responsibility is reconciled, never picked, so a disagreement is flagged instead of resolved."""
    base = controls or pids
    c_resp = controls.get("Responsibility") if controls else None
    p_resp = pids.get("Responsibility") if pids else None
    c_resp = "[unclear]" if _is_unclear(c_resp) else str(c_resp)
    p_resp = "[unclear]" if _is_unclear(p_resp) else str(p_resp)
    conflict = controls is not None and pids is not None and "[unclear]" not in (c_resp, p_resp) and c_resp.upper() != p_resp.upper()

    merged = {"name": base["name"]}
    for field in ("description", "wire_count", "wire_type", "voltage"):
        value = controls.get(field) if controls else None
        merged[field] = "[unclear]" if _is_unclear(value) else value
    merged["intrinsic_safety"] = bool(controls.get("intrinsic_safety")) if controls else False

    if conflict:
        merged["Responsibility"] = "CONFLICT"
    else:
        known = [r for r in (c_resp, p_resp) if r != "[unclear]"]
        merged["Responsibility"] = known[0] if known else "[unclear]"
    merged["responsibility_conflict"] = conflict
    merged["controls_responsibility"] = c_resp if controls else None
    merged["pids_responsibility"] = p_resp if pids else None

    controls_notes = (controls or {}).get("Controls Notes") or (controls or {}).get("Notes")
    pids_notes = (pids or {}).get("P&ID Notes") or (pids or {}).get("Notes")
    merged["controls_notes"] = "" if _is_unclear(controls_notes) else str(controls_notes)
    merged["pids_notes"] = "" if _is_unclear(pids_notes) else str(pids_notes)

    conflicts: List[str] = []
    if conflict:
        conflicts.append(f"RESPONSIBILITY CONFLICT: controls says {c_resp}, P&ID says {p_resp}; verify against the drawings")
    if controls and pids and _tag_keys(controls["name"])[0] != _tag_keys(pids["name"])[0]:  # matched through an alias or a fuzzy pairing
        conflicts.append(f"Tag differs: controls {controls['name']}, P&ID {pids['name']}")
    merged["conflict_notes"] = "; ".join(conflicts)
    merged["source"] = "both" if controls and pids else ("controls" if controls else "pids")
    merged["controls_pages"] = (controls or {}).get("_pages", [])
    merged["pids_pages"] = (pids or {}).get("_pages", [])
    return merged


def combine_stage_outputs(
    controls_output: dict,
    pids_output: dict,
    model: Optional[str] = None,
    api_host: Optional[str] = None,
    api_key: Optional[str] = None,
    max_tokens: int = 2048,
    log_callback: Optional[Callable[[str], None]] = None,
    stream_callback: Optional[Callable[[str], None]] = None,
) -> str:
    """Merge controls and P&ID outputs in code (normalized-tag join); the model only pairs near-miss tags."""
    resolved_model = model or os.getenv("AI_GATEWAY_MODEL", "gemma4:e2b")
    resolved_host = api_host or os.getenv("AI_GATEWAY_HOST", "http://localhost:11434")
    resolved_api_key = api_key or os.getenv("AI_GATEWAY_API_KEY") or "None"

    def _log(message: str) -> None:
        if log_callback is not None:
            log_callback(message)

    controls = _collapse_by_key(_flatten_stage(controls_output))
    pids = _collapse_by_key(_flatten_stage(pids_output))

    excluded: List[dict] = []
    for side, records in (("controls", controls), ("pids", pids)):
        for key in [k for k, r in records.items() if _is_non_instrument(r["name"])]:
            record = records.pop(key)
            excluded.append({"name": record["name"], "source": side, "reason": _is_non_instrument(record["name"])})

    # Exact join on any shared normalized key (handles PT-1207 vs PT1207 and 'SV1210 (XV1210)').
    pid_by_key = {k: pkey for pkey, p in pids.items() for k in p["_keys"]}
    pairs: List[tuple] = []
    for ckey in list(controls):
        pkey = next((pid_by_key[k] for k in controls[ckey]["_keys"] if k in pid_by_key and pid_by_key[k] in pids), None)
        if pkey is not None:
            pairs.append((controls.pop(ckey), pids.pop(pkey)))
    _log(f"Combine: {len(pairs)} tags matched exactly after normalization; "
         f"{len(controls)} controls-only and {len(pids)} P&ID-only remain")

    # Near-miss candidates (same loop number, or very similar text) go to the model as a short pairing question.
    import difflib
    candidates: List[tuple] = []
    for c in controls.values():
        c_key = c["_keys"][0]
        c_digits = re.sub(r"\D", "", c_key)
        for p in pids.values():
            p_key = p["_keys"][0]
            if (c_digits and c_digits == re.sub(r"\D", "", p_key)) or difflib.SequenceMatcher(None, c_key, p_key).ratio() >= 0.8:
                candidates.append((c, p))
    if candidates:
        _log(f"Combine: asking the model to pair {len(candidates)} near-miss candidate(s)")
        for c, p in _ask_model_to_pair(candidates, resolved_model, resolved_host, resolved_api_key, max_tokens,
                                       log_callback, stream_callback):
            ckey, pkey = c["_keys"][0], p["_keys"][0]
            if ckey in controls and pkey in pids:
                pairs.append((controls.pop(ckey), pids.pop(pkey)))
    else:
        _log("Combine: no near-miss candidates; the model is not needed")

    combined = [_merge_pair(c, p) for c, p in pairs]
    combined += [_merge_pair(c, None) for c in controls.values()]
    combined += [_merge_pair(None, p) for p in pids.values()]
    combined.sort(key=lambda r: (not r["responsibility_conflict"], r["name"]))  # conflicts first
    conflicts = [r for r in combined if r["responsibility_conflict"]]

    counts = {s: sum(1 for r in combined if r["source"] == s) for s in ("both", "controls", "pids")}
    result = {
        "plc": controls_output.get("plc") if not _is_unclear(controls_output.get("plc")) else pids_output.get("plc", "[unclear]"),
        "responsibility_conflicts": [
            {"name": r["name"], "controls": r["controls_responsibility"], "pids": r["pids_responsibility"]} for r in conflicts
        ],
        "combined_instruments": combined,
        "excluded": excluded,
        "summary": (f"{len(combined)} instruments: {counts['both']} on both drawings, {counts['controls']} controls only, "
                    f"{counts['pids']} P&ID only; {len(excluded)} non-instrument rows excluded. "
                    f"{len(conflicts)} RESPONSIBILITY CONFLICT(S) between controls and P&ID need review. "
                    "Only tag and responsibility are taken from the P&ID; all other fields come from controls."),
    }
    _log(f"Combine result: {result['summary']}")
    for r in conflicts:
        _log(f"[RESPONSIBILITY CONFLICT] {r['name']}: controls={r['controls_responsibility']} vs P&ID={r['pids_responsibility']}")
    return json.dumps(result, indent=2, ensure_ascii=False)


GUI_INPUTS_FILE = Path(__file__).with_name("instrument_list_gui_inputs.txt")


def load_gui_inputs() -> dict:
    """Load GUI values entered during the previous run."""
    try:
        saved_values = json.loads(GUI_INPUTS_FILE.read_text(encoding="utf-8"))
        if not isinstance(saved_values, dict):
            return {}
        # Sanitize: if output field was corrupted to a provider name, clear it
        output_val = saved_values.get("output", "")
        if isinstance(output_val, str) and output_val.strip() in {"Local", "Anthropic", "OpenAI", "Google", "Custom"}:
            saved_values["output"] = ""
        return saved_values
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_gui_inputs(values: dict) -> None:
    """Save GUI values, including an explicitly entered API key, for the next run."""
    GUI_INPUTS_FILE.write_text(json.dumps(values, indent=2), encoding="utf-8")


def _get_local_ollama_models(api_host: str = "http://localhost:11434") -> List[str]:
    """Return installed Ollama model names when the local endpoint is reachable."""
    try:
        response = requests.get(f"{_normalize_host(api_host, 'http://localhost:11434')}/api/tags", timeout=5)
        response.raise_for_status()
        payload = response.json()
        models: List[str] = []
        for model in payload.get("models", []):
            if isinstance(model, dict):
                name = model.get("name")
                if name:
                    models.append(name)
            elif isinstance(model, str):
                models.append(model)
        return sorted(dict.fromkeys(models))
    except Exception:
        return []


def _get_claude_models(api_key: str, api_host: str = "https://api.anthropic.com") -> List[str]:
    """Query Anthropic for available Claude models using the provided API key."""
    if not api_key or api_key == "None":
        return []

    try:
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
        response = requests.get(
            f"{_normalize_host(api_host, 'https://api.anthropic.com')}/v1/models",
            headers=headers,
            timeout=15,
        )
        response.raise_for_status()
        payload = response.json()
        models: List[str] = []
        for item in payload.get("data", []):
            if isinstance(item, dict):
                model_id = item.get("id")
                if model_id:
                    models.append(model_id)
        return sorted(dict.fromkeys(models))
    except Exception:
        return []


def _get_cloud_models(api_key: str, api_host: str = "", provider: Optional[str] = None) -> List[str]:
    """Discover cloud models from the selected provider when possible."""
    normalized_provider = _normalize_cloud_provider(provider)
    if normalized_provider == "Anthropic":
        return _get_claude_models(api_key, api_host or "https://api.anthropic.com")

    if normalized_provider == "OpenAI":
        if not api_key or api_key == "None":
            return []
        try:
            response = requests.get(
                f"{_normalize_host(api_host or 'https://api.openai.com', 'https://api.openai.com')}/v1/models",
                headers={"Authorization": f"Bearer {api_key}"},
                timeout=15,
            )
            response.raise_for_status()
            payload = response.json()
            models: List[str] = []
            for item in payload.get("data", []):
                if isinstance(item, dict):
                    model_id = item.get("id")
                    if isinstance(model_id, str) and model_id:
                        models.append(model_id)
            return sorted(dict.fromkeys(models))
        except Exception:
            return []

    if normalized_provider == "Google":
        if not api_key or api_key == "None":
            return []
        try:
            response = requests.get(
                f"{_normalize_host(api_host or 'https://generativelanguage.googleapis.com', 'https://generativelanguage.googleapis.com')}/v1/models",
                params={"key": api_key},
                timeout=15,
            )
            response.raise_for_status()
            payload = response.json()
            models: List[str] = []
            for item in payload.get("models", []):
                if isinstance(item, dict):
                    name = item.get("name")
                    if isinstance(name, str):
                        models.append(name.split("/")[-1])
            return sorted(dict.fromkeys(models))
        except Exception:
            return []

    return []


# MODEL_MAX_TOKENS = {
#     "claude-3-5-sonnet-latest": 200000,
#     "claude-sonnet-4-20250514": 200000,
#     "claude-3-7-sonnet": 200000,
#     "claude-3-haiku": 200000,
#     "gemma4:e2b": 4096,
#     "llama3.2:latest": 4096,
#     "mistral": 4096,
#     "qwen": 4096,
#     "phi": 4096,
#     "minicpm-v4.6": 2048,
# }


# def _estimate_max_tokens(model_name: Optional[str], use_local: bool) -> int:
#     """Use a model-capacity map whenever possible, otherwise fall back to a safe default."""
#     name = (model_name or "").strip().lower()
#     if not name:
#         return 2048

#     for candidate, max_tokens in MODEL_MAX_TOKENS.items():
#         if candidate.lower() in name:
#             return max_tokens

#     if "claude" in name or "anthropic" in name:
#         return 200000
#     if "gemma" in name or "llama" in name or "qwen" in name or "mistral" in name or "phi" in name:
#         return 4096
#     if "minicpm" in name:
#         return 2048
#     if use_local:
#         return 2048
#     return 2048


def _get_model_capabilities(model_name: Optional[str], api_host: str, api_key: str, use_local: bool) -> dict:
    """Try to fetch dynamic capability metadata for the selected model and fall back to a safe estimate."""
    resolved_model = (model_name or "").strip()
    # fallback_max_tokens = 4096 if use_local else 0
    # capabilities = {"max_tokens": fallback_max_tokens}
    capabilities = {"max_tokens": 1000000}

    def _coerce_positive_int(value: object) -> Optional[int]:
        if isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            parsed = int(value)
            return parsed if parsed > 0 else None
        if isinstance(value, str):
            match = re.search(r"\b(\d{3,})\b", value)
            if match:
                parsed = int(match.group(1))
                return parsed if parsed > 0 else None
        return None

    def _collect_candidates(payload: object, out: List[int]) -> None:
        if isinstance(payload, dict):
            for key, value in payload.items():
                lowered = str(key).lower()
                if any(token in lowered for token in ("context_length", "context_window", "max_context", "num_ctx", "n_ctx", "max_tokens", "max_output_tokens")):
                    if "embedding" not in lowered:
                        parsed = _coerce_positive_int(value)
                        if parsed is not None:
                            out.append(parsed)
                if isinstance(value, (dict, list)):
                    _collect_candidates(value, out)
                elif isinstance(value, str) and lowered in {"parameters", "parameter", "options", "modelfile"}:
                    for token_match in re.finditer(r"(?:num_ctx|n_ctx|max_tokens)\s*[=:]?\s*(\d+)", value, flags=re.IGNORECASE):
                        out.append(int(token_match.group(1)))
        elif isinstance(payload, list):
            for item in payload:
                _collect_candidates(item, out)

    def _update_from_payload(payload: object) -> None:
        if not isinstance(payload, dict):
            return
        candidates: List[int] = []
        _collect_candidates(payload, candidates)
        if candidates:
            # Keep the largest discovered context/token capacity.
            capabilities["max_tokens"] = max(candidates)

    if use_local:
        try:
            response = requests.post(
                f"{_normalize_host(api_host, 'http://localhost:11434')}/api/show",
                headers={"Content-Type": "application/json"},
                data=json.dumps({"name": resolved_model}),
                timeout=10,
            )
            response.raise_for_status()
            payload = response.json()
            _update_from_payload(payload)
        except Exception:
            pass
        return capabilities

    if not _is_claude_model(resolved_model):
        return capabilities

    try:
        headers = {
            "x-api-key": api_key or "",
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
        response = requests.get(
            f"{_normalize_host(api_host, 'https://api.anthropic.com')}/v1/models",
            headers=headers,
            timeout=15,
        )
        response.raise_for_status()
        payload = response.json()
        _update_from_payload(payload)
        if isinstance(payload.get("data"), list):
            for item in payload["data"]:
                if not isinstance(item, dict):
                    continue
                if item.get("id") != resolved_model:
                    continue
                _update_from_payload(item)
                break
    except Exception:
        pass

    return capabilities


def _get_pdf_native_dpi(pdf_path: Path) -> int:
    """Estimate a sensible native render DPI from embedded raster images in the PDF."""
    if fitz is None:
        return 300
    try:
        doc = fitz.open(str(pdf_path))
        if not doc.page_count:
            return 300

        inferred_dpis: List[int] = []
        for page in doc:
            try:
                for image_info in page.get_image_info(xrefs=True):
                    width = image_info.get("width") or 0
                    height = image_info.get("height") or 0
                    if width and height:
                        inferred_dpis.append(max(width, height))
            except Exception:
                continue

        if inferred_dpis:
            max_pixels = max(inferred_dpis)
            if max_pixels >= 4000:
                return 600
            if max_pixels >= 2000:
                return 400
            return 300

        return 300
    except Exception:
        return 300


def _resolve_model_settings(
    model_name: Optional[str],
    use_local: bool,
    current_api_host: str,
    current_api_key: str,
    cloud_provider: Optional[str] = None,
) -> tuple[str, str, str, int]:
    """Compute the host, API key, and token limit for the selected model."""
    resolved_model = (model_name or "").strip() or os.getenv("AI_GATEWAY_MODEL", "gemma4:e2b")
    normalized_provider = _normalize_cloud_provider(cloud_provider)
    if normalized_provider == "Local":
        resolved_host = "http://localhost:11434"
        resolved_api_key = ""
    elif normalized_provider == "OpenAI":
        resolved_host = current_api_host or os.getenv("AI_GATEWAY_HOST", "https://api.openai.com")
        resolved_api_key = current_api_key or os.getenv("AI_GATEWAY_API_KEY", "")
    elif normalized_provider == "Google":
        resolved_host = current_api_host or os.getenv("AI_GATEWAY_HOST", "https://generativelanguage.googleapis.com")
        resolved_api_key = current_api_key or os.getenv("AI_GATEWAY_API_KEY", "")
    elif _is_claude_model(resolved_model):
        resolved_host = "https://api.anthropic.com"
        resolved_api_key = current_api_key or os.getenv("AI_GATEWAY_API_KEY", "")
    else:
        resolved_host = current_api_host or os.getenv("AI_GATEWAY_HOST", "http://localhost:11434")
        resolved_api_key = current_api_key or os.getenv("AI_GATEWAY_API_KEY", "")

    if normalized_provider != "Local":
        discovered_models = _get_cloud_models(resolved_api_key, resolved_host, normalized_provider)
        if discovered_models and resolved_model not in discovered_models:
            resolved_model = discovered_models[0]

    capabilities = _get_model_capabilities(resolved_model, resolved_host, resolved_api_key, normalized_provider == "Local")
    resolved_max_tokens = int(capabilities.get("max_tokens") or 0)
    return resolved_model, resolved_host, resolved_api_key, resolved_max_tokens


def run_pipeline(
    parsed_args: argparse.Namespace,
    log_callback: Callable[[str], None],
    stream_callback: Optional[Callable[[str], None]] = None,
) -> int:
    """Run the stage(s) selected by parsed_args.run_mode, logging through log_callback. Used by the GUI."""
    def _log(msg: str) -> None:
        log_callback(msg)

    run_mode = parsed_args.run_mode or "all"

    controls_payload: Optional[dict] = None
    pids_payload: Optional[dict] = None
    fallback_pdf = Path(parsed_args.pdf_path).expanduser().resolve() if parsed_args.pdf_path else None

    if run_mode in {"all", "controls"}:
        controls_pdf = Path(parsed_args.controls_pdf).expanduser().resolve() if parsed_args.controls_pdf else fallback_pdf
        if controls_pdf is None or not controls_pdf.exists():
            raise ValueError("Controls stage requires --controls-pdf (or fallback pdf_path).")
        controls_pages = _resolve_pages_for_pdf(controls_pdf, parsed_args.controls_pages)
        _log(f"Running controls stage on {controls_pdf}")
        controls_json = transcribe_controls_drawings(
            pdf_path=controls_pdf,
            model=parsed_args.controls_model or parsed_args.model,
            api_host=parsed_args.controls_api_host or parsed_args.api_host,
            api_key=parsed_args.controls_api_key or parsed_args.api_key,
            max_tokens=parsed_args.controls_max_tokens if parsed_args.controls_max_tokens is not None else parsed_args.max_tokens,
            dpi=parsed_args.controls_dpi if parsed_args.controls_dpi is not None else parsed_args.dpi,
            pages=controls_pages,
            reference_image=Path(parsed_args.controls_reference_image).expanduser().resolve() if parsed_args.controls_reference_image else (Path(parsed_args.reference_image).expanduser().resolve() if parsed_args.reference_image else None),
            log_callback=_log,
            stream_callback=stream_callback,
            docling_enable=parsed_args.docling_enable,
            docling_do_ocr=parsed_args.docling_do_ocr,
            docling_do_table_structure=parsed_args.docling_do_table_structure,
            docling_do_picture_description=parsed_args.docling_do_picture_description,
            docling_images_scale=parsed_args.docling_images_scale,
        )
        controls_payload = _extract_json_object(controls_json)
        if parsed_args.controls_output:
            _log("Saving controls extraction results...")
            _save_json_text(controls_json, parsed_args.controls_output, log_callback=_log)

    if run_mode in {"all", "pids"}:
        pids_pdf = Path(parsed_args.pids_pdf).expanduser().resolve() if parsed_args.pids_pdf else fallback_pdf
        if pids_pdf is None or not pids_pdf.exists():
            raise ValueError("P&ID stage requires --pids-pdf (or fallback pdf_path).")
        pids_pages = _resolve_pages_for_pdf(pids_pdf, parsed_args.pids_pages)
        _log(f"Running P&ID stage on {pids_pdf}")
        pids_json = transcribe_pids_drawings(
            pdf_path=pids_pdf,
            model=parsed_args.pids_model or parsed_args.model,
            api_host=parsed_args.pids_api_host or parsed_args.api_host,
            api_key=parsed_args.pids_api_key or parsed_args.api_key,
            max_tokens=parsed_args.pids_max_tokens if parsed_args.pids_max_tokens is not None else parsed_args.max_tokens,
            dpi=parsed_args.pids_dpi if parsed_args.pids_dpi is not None else parsed_args.dpi,
            pages=pids_pages,
            reference_image=Path(parsed_args.pids_reference_image).expanduser().resolve() if parsed_args.pids_reference_image else (Path(parsed_args.reference_image).expanduser().resolve() if parsed_args.reference_image else None),
            log_callback=_log,
            stream_callback=stream_callback,
            docling_enable=parsed_args.docling_enable,
            docling_do_ocr=parsed_args.docling_do_ocr,
            docling_do_table_structure=parsed_args.docling_do_table_structure,
            docling_do_picture_description=parsed_args.docling_do_picture_description,
            docling_images_scale=parsed_args.docling_images_scale,
        )
        pids_payload = _extract_json_object(pids_json)
        if parsed_args.pids_output:
            _log("Saving P&ID extraction results...")
            _save_json_text(pids_json, parsed_args.pids_output, log_callback=_log)

    if run_mode in {"all", "combine"}:
        if controls_payload is None:
            if not parsed_args.controls_input_json:
                raise ValueError("Combination stage requires controls JSON input.")
            controls_payload = _load_json_dict(Path(parsed_args.controls_input_json).expanduser().resolve())
        if pids_payload is None:
            if not parsed_args.pids_input_json:
                raise ValueError("Combination stage requires P&ID JSON input.")
            pids_payload = _load_json_dict(Path(parsed_args.pids_input_json).expanduser().resolve())

        _log("Running combination stage...")
        combined_json = combine_stage_outputs(
            controls_output=controls_payload,
            pids_output=pids_payload,
            model=parsed_args.combine_model or parsed_args.model,
            api_host=parsed_args.combine_api_host or parsed_args.api_host,
            api_key=parsed_args.combine_api_key or parsed_args.api_key,
            max_tokens=parsed_args.combine_max_tokens if parsed_args.combine_max_tokens is not None else parsed_args.max_tokens,
            log_callback=_log,
            stream_callback=stream_callback,
        )
        combine_out = parsed_args.combine_output or parsed_args.output
        if combine_out:
            _log("Saving combined analysis results...")
            _save_json_text(combined_json, combine_out, log_callback=_log)

    return 0


def get_gui_arguments(parser: argparse.ArgumentParser) -> Optional[argparse.Namespace]:
    """Open the GUI (see Instrument_List_GUI.py); it runs the pipeline itself, so there is nothing to return."""
    from Instrument_List_GUI import run_gui
    run_gui(parser)
    return None


def _load_json_dict(json_path: Path) -> dict:
    payload = json.loads(json_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError(f"Expected a top-level JSON object in {json_path}")
    return payload


def _save_json_text(json_text: str, output_path: Optional[str], log_callback: Optional[Callable[[str], None]] = None) -> None:
    if not output_path:
        return
    path = Path(output_path).expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json_text, encoding="utf-8")
    file_size = len(json_text)
    if log_callback:
        log_callback(f"Saved to: {path}")
        log_callback(f"File size: {file_size} bytes")
    print(f"Saved output to {path}")


def _resolve_pages_for_pdf(pdf_path: Path, pages_text: Optional[str]) -> Optional[List[int]]:
    if not pages_text or not pages_text.strip():
        return None
    if fitz is None:
        raise RuntimeError("PyMuPDF is required to parse page selections. Install it with: pip install pymupdf")
    doc = fitz.open(str(pdf_path))
    total_pages = len(doc)
    doc.close()
    selected_pages = parse_page_selection(pages_text, total_pages)
    if not selected_pages:
        raise ValueError(f"Page selection '{pages_text}' matched no pages in this {total_pages}-page PDF.")
    return selected_pages


def _cli_stream(text: str) -> None:
    """Live model output for command-line runs; stderr keeps stdout clean for the final JSON."""
    print(text, end="", file=sys.stderr, flush=True)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Extract instruments from controls and P&ID drawings, then optionally combine both analyses")
    parser.add_argument("pdf_path", nargs="?", help="Fallback PDF path for the controls/P&ID stages when --controls-pdf / --pids-pdf is not given")
    parser.add_argument("--run-mode", choices=["all", "controls", "pids", "combine"], default="all", help="Select which stage(s) to run")

    # Legacy/global options
    parser.add_argument("--model", help="AI model name (for example: gemma4:e2b or claude-3-5-sonnet-latest)")
    parser.add_argument("--api-host", help="API host for the model endpoint")
    parser.add_argument("--api-key", help="API key for the model endpoint")
    parser.add_argument("--max-tokens", type=int, default=2048, help="Maximum tokens to generate")
    parser.add_argument("--dpi", type=int, default=300, help="Rendering DPI for each PDF page")
    parser.add_argument("--output", help="Optional path to save the JSON transcription")
    parser.add_argument("--reference-image", help="Optional reference image sent to the AI alongside each page")

    # Controls stage options
    parser.add_argument("--controls-pdf", help="PDF path for controls extraction stage")
    parser.add_argument("--controls-pages", help="Controls page selection (for example: 1,3-5)")
    parser.add_argument("--controls-model", help="Model for controls extraction stage")
    parser.add_argument("--controls-api-host", help="API host for controls extraction stage")
    parser.add_argument("--controls-api-key", help="API key for controls extraction stage")
    parser.add_argument("--controls-max-tokens", type=int, help="Max tokens for controls extraction stage")
    parser.add_argument("--controls-dpi", type=int, help="Render DPI for controls extraction stage")
    parser.add_argument("--controls-reference-image", help="Optional reference image for controls extraction")
    parser.add_argument("--controls-output", help="Optional output JSON path for controls extraction")

    # P&ID stage options
    parser.add_argument("--pids-pdf", help="PDF path for P&ID extraction stage")
    parser.add_argument("--pids-pages", help="P&ID page selection (for example: 1,3-5)")
    parser.add_argument("--pids-model", help="Model for P&ID extraction stage")
    parser.add_argument("--pids-api-host", help="API host for P&ID extraction stage")
    parser.add_argument("--pids-api-key", help="API key for P&ID extraction stage")
    parser.add_argument("--pids-max-tokens", type=int, help="Max tokens for P&ID extraction stage")
    parser.add_argument("--pids-dpi", type=int, help="Render DPI for P&ID extraction stage")
    parser.add_argument("--pids-reference-image", help="Optional reference image for P&ID extraction")
    parser.add_argument("--pids-output", help="Optional output JSON path for P&ID extraction")

    # Combination stage options
    parser.add_argument("--controls-input-json", help="Controls JSON input path for combine-only mode")
    parser.add_argument("--pids-input-json", help="P&ID JSON input path for combine-only mode")
    parser.add_argument("--combine-model", help="Model for combination stage")
    parser.add_argument("--combine-api-host", help="API host for combination stage")
    parser.add_argument("--combine-api-key", help="API key for combination stage")
    parser.add_argument("--combine-max-tokens", type=int, help="Max tokens for combination stage")
    parser.add_argument("--combine-output", help="Optional output JSON path for combination stage")
    return parser


def main() -> int:

    print("Starting main function")
    print("Parsing command-line arguments")
    parser = build_arg_parser()
    print("Finished setting up argument parser")

    if len(sys.argv) == 1:
        args = get_gui_arguments(parser)
        if args is None:
            return 0
    else:
        args = parser.parse_args()
    print("Arguments: ", args)
    run_mode = args.run_mode or "all"

    controls_payload: Optional[dict] = None
    pids_payload: Optional[dict] = None

    try:
        fallback_pdf = Path(args.pdf_path).expanduser().resolve() if args.pdf_path else None

        if run_mode in {"all", "controls"}:
            controls_pdf = Path(args.controls_pdf).expanduser().resolve() if args.controls_pdf else fallback_pdf
            if controls_pdf is None or not controls_pdf.exists():
                raise ValueError("Controls stage requires --controls-pdf (or the positional pdf_path fallback).")
            controls_pages = _resolve_pages_for_pdf(controls_pdf, args.controls_pages)
            controls_json = transcribe_controls_drawings(
                pdf_path=controls_pdf,
                stream_callback=_cli_stream,
                model=args.controls_model or args.model,
                api_host=args.controls_api_host or args.api_host,
                api_key=args.controls_api_key or args.api_key,
                max_tokens=args.controls_max_tokens if args.controls_max_tokens is not None else args.max_tokens,
                dpi=args.controls_dpi if args.controls_dpi is not None else args.dpi,
                pages=controls_pages,
                reference_image=Path(args.controls_reference_image).expanduser().resolve() if args.controls_reference_image else (Path(args.reference_image).expanduser().resolve() if args.reference_image else None),
            )
            controls_payload = _extract_json_object(controls_json)
            print("=" * 30 + " CONTROLS OUTPUT " + "=" * 30)
            print(controls_json)
            _save_json_text(controls_json, args.controls_output)

        if run_mode in {"all", "pids"}:
            pids_pdf = Path(args.pids_pdf).expanduser().resolve() if args.pids_pdf else fallback_pdf
            if pids_pdf is None or not pids_pdf.exists():
                raise ValueError("P&ID stage requires --pids-pdf (or the positional pdf_path fallback).")
            pids_pages = _resolve_pages_for_pdf(pids_pdf, args.pids_pages)
            pids_json = transcribe_pids_drawings(
                pdf_path=pids_pdf,
                stream_callback=_cli_stream,
                model=args.pids_model or args.model,
                api_host=args.pids_api_host or args.api_host,
                api_key=args.pids_api_key or args.api_key,
                max_tokens=args.pids_max_tokens if args.pids_max_tokens is not None else args.max_tokens,
                dpi=args.pids_dpi if args.pids_dpi is not None else args.dpi,
                pages=pids_pages,
                reference_image=Path(args.pids_reference_image).expanduser().resolve() if args.pids_reference_image else (Path(args.reference_image).expanduser().resolve() if args.reference_image else None),
            )
            pids_payload = _extract_json_object(pids_json)
            print("=" * 31 + " P&ID OUTPUT " + "=" * 31)
            print(pids_json)
            _save_json_text(pids_json, args.pids_output)

        if run_mode in {"all", "combine"}:
            if controls_payload is None:
                if not args.controls_input_json:
                    raise ValueError("Combination stage requires controls JSON from --controls-input-json when controls stage was not run.")
                controls_payload = _load_json_dict(Path(args.controls_input_json).expanduser().resolve())
            if pids_payload is None:
                if not args.pids_input_json:
                    raise ValueError("Combination stage requires P&ID JSON from --pids-input-json when P&ID stage was not run.")
                pids_payload = _load_json_dict(Path(args.pids_input_json).expanduser().resolve())

            combined_json = combine_stage_outputs(
                controls_output=controls_payload,
                pids_output=pids_payload,
                stream_callback=_cli_stream,
                model=args.combine_model or args.model,
                api_host=args.combine_api_host or args.api_host,
                api_key=args.combine_api_key or args.api_key,
                max_tokens=args.combine_max_tokens if args.combine_max_tokens is not None else args.max_tokens,
            )
            print("=" * 28 + " COMBINATION OUTPUT " + "=" * 28)
            print(combined_json)
            _save_json_text(combined_json, args.combine_output or args.output)

    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    main()
