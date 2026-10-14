"""PDF pre-step shared by both tools, mirroring the Document_To_Database_Pipeline_Ollama preprocessing.

For every PDF input:
  1. clean the PDF (ocrmypdf deskew + upsample, as in Vector_Database/Clean_PDF_For_OCR.py), falling back to
     the original file if cleaning isn't possible;
  2. convert each page with Docling using full-page OCR (the embedded text layer is thrown away and every
     page is OCR'd, as in Vector_Database/Docling_Convert_Ollama.py) and keep the markdown per page.

Results are cached on disk keyed by the PDF and the OCR settings.

Environment switches (same names as the Vector_Database pipeline):
    PIPELINE_CLEAN_PDF      1/0 (default 1)  run the ocrmypdf cleaning step
    DOCLING_OCR_ENGINE      rapidocr (default here) | easyocr | tesseract | auto
    DOCLING_FULL_PAGE_OCR   1/0 (default 1)  OCR whole pages instead of only bitmap regions
"""
import hashlib
import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Callable, Dict, List, Optional

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from docling.datamodel.accelerator_options import AcceleratorDevice, AcceleratorOptions
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import (
        EasyOcrOptions,
        OcrAutoOptions,
        PdfPipelineOptions,
        RapidOcrOptions,
        TesseractCliOcrOptions,
    )
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling_core.types.doc.base import ImageRefMode
    DOCLING_AVAILABLE = True
except ImportError:
    DOCLING_AVAILABLE = False

try:  # newer Docling selects full-page OCR with `mode`; older releases use `force_full_page_ocr`
    from docling.datamodel.pipeline_options import OcrMode
except ImportError:
    OcrMode = None

OCR_CACHE_DIR = Path(__file__).with_name("ocr_cache")
CLEAN_PDF_DEEP = False  # True also runs unpaper cleanup (needs unpaper on PATH)


def _env_on(name: str, default: str = "1") -> bool:
    return os.getenv(name, default).strip().lower() in ("1", "true", "yes", "on")


def clean_pdf(input_path: Path, output_path: Path, oversample: int = 300, extra_clean: bool = False) -> None:
    """Deskew/upsample a PDF with ocrmypdf (its own OCR text is discarded; Docling does the real OCR)."""
    if shutil.which("tesseract") is None:
        raise RuntimeError("tesseract not found on PATH (ocrmypdf needs an OCR engine present to run)")
    args = [
        sys.executable, "-m", "ocrmypdf",
        "--deskew",
        "--oversample", str(oversample),
        "--force-ocr",  # rasterize + reprocess every page, even ones with an existing (bad) text layer
        "--output-type", "pdf",
        "--optimize", "2" if shutil.which("pngquant") else "1",
        "--jpeg-quality", "85",
    ]
    if extra_clean:
        if shutil.which("unpaper") is None:
            raise RuntimeError("deep cleaning requires unpaper on PATH")
        args += ["--clean", "--clean-final"]
    args += [str(input_path), str(output_path)]
    result = subprocess.run(args, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"ocrmypdf exited with code {result.returncode}: {result.stderr.strip()[-500:]}")


def clean_pdf_step(pdf_path: Path, cleaned_path: Path, log: Callable[[str], None]) -> Path:
    """Return a cleaned copy of the PDF, reusing an earlier one; fall back to the original on any failure."""
    if not _env_on("PIPELINE_CLEAN_PDF"):
        return pdf_path
    if cleaned_path.is_file() and cleaned_path.stat().st_mtime >= pdf_path.stat().st_mtime:
        log(f"Using previously cleaned PDF: {cleaned_path.name}")
        return cleaned_path
    cleaned_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        log(f"Cleaning PDF (deskew/upsample): {pdf_path.name}")
        clean_pdf(pdf_path, cleaned_path, extra_clean=CLEAN_PDF_DEEP)
        return cleaned_path
    except Exception as exc:
        log(f"PDF cleaning failed, using the original file instead ({exc})")
        cleaned_path.unlink(missing_ok=True)
        return pdf_path


def build_ocr_options(engine: Optional[str] = None, full_page: Optional[bool] = None):
    """Docling OCR options; full-page OCR throws away the embedded text layer and OCRs every page."""
    engine = (engine or os.getenv("DOCLING_OCR_ENGINE", "rapidocr")).strip().lower()
    if full_page is None:
        full_page = _env_on("DOCLING_FULL_PAGE_OCR")

    if OcrMode is not None:
        mode = {"mode": OcrMode.FULL_PAGE if full_page else OcrMode.DEFAULT}
    else:
        mode = {"force_full_page_ocr": full_page}

    if engine == "rapidocr":
        return RapidOcrOptions(bitmap_area_threshold=0.001, **mode)
    if engine == "easyocr":
        return EasyOcrOptions(lang=["en"], use_gpu=True, **mode)
    if engine == "tesseract":
        return TesseractCliOcrOptions(lang=["eng"], **mode)
    if engine != "auto":
        raise ValueError(f"Unknown DOCLING_OCR_ENGINE '{engine}' (use auto, rapidocr, easyocr or tesseract)")
    return OcrAutoOptions(**mode)


def _cache_path(pdf_path: Path, do_table_structure: bool, images_scale: float) -> Path:
    """Cache file keyed by the PDF's identity (path, size, mtime) and everything that changes the OCR text."""
    stat = pdf_path.stat()
    key = "|".join(str(v) for v in (
        pdf_path.resolve(), stat.st_size, stat.st_mtime_ns, do_table_structure, images_scale,
        _env_on("PIPELINE_CLEAN_PDF"), os.getenv("DOCLING_OCR_ENGINE", "rapidocr"), _env_on("DOCLING_FULL_PAGE_OCR"),
    ))
    return OCR_CACHE_DIR / f"{pdf_path.stem}_{hashlib.sha1(key.encode()).hexdigest()[:12]}.json"


def preprocess_pdf_ocr(
    pdf_path: Path,
    pages: Optional[List[int]] = None,
    do_table_structure: bool = True,
    images_scale: float = 3.0,
    log_callback: Optional[Callable[[str], None]] = None,
) -> Dict[int, str]:
    """Pre-step for every PDF input: clean it, then full-page OCR each page into markdown.

    Returns {page_number: markdown}; pages that fail or have no text are omitted. Pages are converted one
    at a time (large drawing sheets otherwise exhaust memory), and results are cached so repeat runs and
    the several stages reading the same PDF only OCR a page once.
    """
    def log(msg: str) -> None:
        if log_callback:
            log_callback(f"[OCR PRE-STEP] {msg}")

    if not DOCLING_AVAILABLE or fitz is None:
        log("Docling or PyMuPDF is not installed; skipping the OCR pre-step")
        return {}

    pdf_path = Path(pdf_path)
    with fitz.open(str(pdf_path)) as doc_pdf:
        page_count = len(doc_pdf)
    wanted = [p for p in (pages or range(1, page_count + 1)) if 1 <= p <= page_count]

    cache_path = _cache_path(pdf_path, do_table_structure, images_scale)
    cached: Dict[int, str] = {}
    if cache_path.is_file():
        try:
            cached = {int(k): v for k, v in json.loads(cache_path.read_text(encoding="utf-8"))["pages"].items()}
        except Exception as exc:
            log(f"Ignoring unreadable cache {cache_path.name}: {exc}")

    missing = [p for p in wanted if p not in cached]
    log(f"{pdf_path.name}: {len(wanted)} page(s) requested, {len(wanted) - len(missing)} cached, {len(missing)} to OCR")
    if not missing:
        return {p: cached[p] for p in wanted if p in cached}

    source = clean_pdf_step(pdf_path, OCR_CACHE_DIR / "_cleaned" / pdf_path.name, log)

    pipeline_options = PdfPipelineOptions(
        ocr_options=build_ocr_options(),
        do_ocr=True,
        do_table_structure=do_table_structure,
        do_picture_description=False,
        images_scale=images_scale,
        accelerator_options=AcceleratorOptions(device=AcceleratorDevice.AUTO),
    )
    converter = DocumentConverter(format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)})

    for page_num in missing:
        log(f"OCR page {page_num}/{page_count}...")
        temp_pdf = Path(tempfile.mktemp(suffix=".pdf"))
        try:
            with fitz.open(str(source)) as src, fitz.open() as single:
                single.insert_pdf(src, from_page=page_num - 1, to_page=page_num - 1)
                single.save(str(temp_pdf))
            markdown = converter.convert(str(temp_pdf)).document.export_to_markdown(
                image_mode=ImageRefMode.REFERENCED
            ).strip()
            if markdown:
                cached[page_num] = markdown
                log(f"Page {page_num}: {len(markdown)} characters")
            else:
                log(f"Page {page_num}: no text recovered")
        except Exception as exc:
            log(f"Page {page_num} failed: {exc}")
        finally:
            temp_pdf.unlink(missing_ok=True)

    try:
        OCR_CACHE_DIR.mkdir(exist_ok=True)
        cache_path.write_text(
            json.dumps({"source": pdf_path.name, "pages": cached}, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as exc:
        log(f"Could not write cache {cache_path}: {exc}")

    return {p: cached[p] for p in wanted if p in cached}
