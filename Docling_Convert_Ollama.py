import os

# Ensure LibreOffice path is set for Docling
os.environ["DOCLING_LIBREOFFICE_CMD"] = r"C:\\Program Files\\LibreOffice\\program\\soffice.exe"
import time
import json
import tempfile
from pathlib import Path
import logging
import pandas as pd
from docling.document_converter import DocumentConverter, PdfFormatOption, WordFormatOption, PowerpointFormatOption
from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import PdfPipelineOptions
from docling.datamodel.settings import settings as docling_settings
from docling.datamodel.accelerator_options import AcceleratorOptions, AcceleratorDevice
from docling_core.types.doc.document import DoclingDocument, PictureItem, TableItem
from docling_core.types.doc.base import ImageRefMode
from pdfconverter.convertword import convert_word
from pdfconverter.convertexcel import convert_excel
from pdfconverter.convertimage import convertimage
import pptxtopdf
import mimetypes
import tempfile
import shutil
import socket
from PIL import Image

# def enable_offline_mode(block_network=True):
#     # HF/Transformers offline flags
#     os.environ["HF_HUB_OFFLINE"] = "1"
#     os.environ["TRANSFORMERS_OFFLINE"] = "1"
#     os.environ["HF_DATASETS_OFFLINE"] = "1"
#     os.environ["HF_HUB_DISABLE_TELEMETRY"] = "1"
#     os.environ["NO_PROXY"] = "*"

#     # # Optional: hard kill all outbound network for this Python process
#     # if block_network:
#     #     def _blocked(*args, **kwargs):
#     #         raise RuntimeError("Network disabled by offline mode")
#     #     socket.create_connection = _blocked
#     #     socket.socket.connect = _blocked

# enable_offline_mode(block_network=True)

# Configure logging (console only)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Use the NVIDIA GPU (CUDA) for model inference (layout, table structure, OCR).
# Requires a CUDA-enabled build of PyTorch to be installed.
accelerator_options = AcceleratorOptions(device=AcceleratorDevice.CUDA)

# Configure pipeline options for PDF files
pdf_pipeline_options = PdfPipelineOptions()
pdf_pipeline_options.accelerator_options = accelerator_options
# NOTE: images_scale controls the DPI at which each PDF page is rasterized
# (scale 1.0 == 72 DPI). High values (e.g. 3.0 == 216 DPI) can exhaust host RAM
# on large documents and trigger `std::bad_alloc` in the preprocess stage. Memory
# grows with the SQUARE of the scale, so large-format engineering drawings can
# still blow up at 1.5. Since page/picture images are not generated for PDFs, a
# scale of 1.0 (72 DPI) is enough to feed OCR while keeping memory bounded.
pdf_pipeline_options.images_scale = 3.0
pdf_pipeline_options.generate_picture_images = True
pdf_pipeline_options.do_table_structure = True
pdf_pipeline_options.generate_table_images = True
pdf_pipeline_options.do_picture_description = True  # Enable picture description generation
pdf_pipeline_options.do_ocr = True  # Enable OCR for image content extraction
# pdf_pipeline_options.do_picture_classification = True  # Enable picture classification to identify image types (e.g., chart, photo, etc.)
# pdf_pipeline_options.enable_remote_services = False  # Ensure no remote calls are made for image processing
pdf_pipeline_options.generate_page_images = False  # Generate full-page images to capture complex layouts and embedded content
pdf_pipeline_options.allow_external_plugins = True  

# Limit how many pages are held in memory / processed per batch. Large PDFs with
# oversized pages (e.g. engineering drawings) can otherwise exhaust host RAM and
# raise `std::bad_alloc` during the preprocess stage. Processing one page at a
# time keeps peak memory to a single rasterized page.
docling_settings.perf.page_batch_size = 1


# Initialize DocumentConverter with format-specific options
doc_converter = DocumentConverter(
    format_options={
        InputFormat.PDF: PdfFormatOption(pipeline_options=pdf_pipeline_options),
        InputFormat.DOCX: WordFormatOption(
            pipeline_options=PdfPipelineOptions(  # Word uses same pipeline options as PDF
                images_scale=3.0,
                generate_picture_images=True,
                do_table_structure=True,
                accelerator_options=accelerator_options,
            )
        ),
        InputFormat.PPTX: PowerpointFormatOption(
            pipeline_options=PdfPipelineOptions(  # PowerPoint uses same pipeline options as PDF
                images_scale=3.0,
                generate_picture_images=True,
                do_table_structure=True,
                accelerator_options=accelerator_options,
            )
        ),
    }
)

# Define supported file extensions for each converter
WORD_EXTENSIONS = {'.doc', '.docx', '.rtf', '.odt'}
EXCEL_EXTENSIONS = {'.xls', '.xlsx', '.ods', '.csv'}
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.webp'}
POWERPOINT_EXTENSIONS = {'.ppt', '.pptx', '.odp'}
PDF_EXTENSIONS = {'.pdf'}

# All supported extensions
ALL_SUPPORTED_EXTENSIONS = (
    WORD_EXTENSIONS | EXCEL_EXTENSIONS | IMAGE_EXTENSIONS | 
    POWERPOINT_EXTENSIONS | PDF_EXTENSIONS
)


def collect_saved_images(image_dir):
    """Return saved image files sorted by name for stable placeholder mapping."""
    if not image_dir.exists():
        return []

    image_files = [
        p for p in image_dir.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS
    ]
    return sorted(image_files, key=lambda p: p.name.lower())


def inject_image_references(markdown_text, image_paths):
    """Replace each <!-- image --> placeholder with a markdown image reference."""
    placeholder = "<!-- image -->"
    if not markdown_text or placeholder not in markdown_text:
        return markdown_text

    parts = markdown_text.split(placeholder)
    if len(parts) == 1:
        return markdown_text

    rebuilt = [parts[0]]
    for idx, part in enumerate(parts[1:], start=1):
        if idx - 1 < len(image_paths):
            image_rel_path = (Path("images") / image_paths[idx - 1].name).as_posix()
            rebuilt.append(f"![Image {idx}]({image_rel_path})")
        else:
            rebuilt.append(placeholder)
        rebuilt.append(part)

    return "".join(rebuilt)


def save_page_images(conv_res, page_image_dir):
    """Save full-page images (when available) to the image folder."""
    pages = getattr(conv_res.document, "pages", None)
    if not pages:
        return 0

    # Support both dict-like and iterable page collections.
    try:
        if hasattr(pages, "items"):
            page_items = sorted(list(pages.items()), key=lambda kv: kv[0])
        else:
            page_items = list(enumerate(pages, start=1))
    except Exception:
        return 0

    saved_count = 0
    for fallback_idx, (page_no, page) in enumerate(page_items, start=1):
        pil_image = None

        try:
            page_image = getattr(page, "image", None)
            if page_image is not None:
                if hasattr(page_image, "pil_image") and page_image.pil_image is not None:
                    pil_image = page_image.pil_image
                elif hasattr(page_image, "to_pil"):
                    pil_image = page_image.to_pil()
                elif hasattr(page_image, "to_pillow"):
                    pil_image = page_image.to_pillow()

            if pil_image is None and hasattr(page, "get_image"):
                try:
                    pil_image = page.get_image(conv_res.document)
                except TypeError:
                    pil_image = page.get_image()
        except Exception as e:
            logging.debug(f"Could not extract page image for page {page_no}: {e}")
            continue

        if pil_image is None:
            continue

        try:
            page_index = int(page_no)
        except Exception:
            page_index = fallback_idx

        page_filename = page_image_dir / f"page_{page_index:04d}.png"
        try:
            page_filename.parent.mkdir(parents=True, exist_ok=True)
            pil_image.save(page_filename)
            saved_count += 1
        except Exception as e:
            logging.debug(f"Could not save page image {page_filename}: {e}")

    if saved_count > 0:
        logging.info(f"Saved {saved_count} full-page images to {page_image_dir}")

    return saved_count


def describe_page_image_with_vision(page_image_path):
    """Use a vision model to transcribe and describe a schematic-like page image."""
    try:
        import ollama

        vision_host = os.getenv("AI_VISION_HOST", "http://localhost:11434")
        candidate_models = [
            os.getenv("AI_VISION_MODEL"),
            "gemma4:12b"
        ]
        candidate_models = [model for model in candidate_models if model]

        image_bytes = page_image_path.read_bytes()
        prompt = (
            "Transcribe the image for the engineering student."
        )

        last_error = None
        for model in candidate_models:
            try:
                client = ollama.Client(host=vision_host)
                response = client.chat(
                    model=model,
                    messages=[
                        {
                            "role": "user",
                            "content": prompt,
                            "images": [image_bytes],
                        }
                    ],
                )

                content = None
                if hasattr(response, "message") and hasattr(response.message, "content"):
                    content = response.message.content
                elif isinstance(response, dict):
                    content = response.get("message", {}).get("content")

                if content and content.strip():
                    return content.strip(), model
            except Exception as e:
                last_error = e
                logging.warning(f"Vision model {model} failed for {page_image_path.name}: {e}")

        if last_error:
            logging.warning(f"No vision description generated for {page_image_path.name}: {last_error}")
        return None, None

    except Exception as e:
        logging.warning(f"Could not run vision description for {page_image_path.name}: {e}")
        return None, None


def _tile_page_image(page_image_path, tile_size=1536, overlap=160):
    """Split a page image into tiles, preferring vision-guided bounds from the full page."""
    try:
        image = Image.open(page_image_path).convert("RGB")
    except Exception as e:
        logging.warning(f"Could not open page image {page_image_path.name}: {e}")
        return []

    width, height = image.size
    if width <= 0 or height <= 0:
        return []

    min_tiles = int(os.getenv("AI_VISION_MIN_TILES", "6"))

    # First try: ask the vision model for JSON tile bounds on the full page.
    vision_tiles = _get_vision_guided_tile_bounds(page_image_path, width, height, min_tiles=min_tiles)
    if vision_tiles:
        tiles = []
        for idx, bbox in enumerate(vision_tiles, start=1):
            x1, y1, x2, y2 = bbox
            tile = image.crop((x1, y1, x2, y2))
            tiles.append({
                "image": tile,
                "bbox": (x1, y1, x2, y2),
                "row": -1,
                "col": idx - 1,
            })
        logging.info(f"Using vision-guided tile plan for {page_image_path.name}: {len(tiles)} tiles")
        return tiles

    logging.warning(f"Falling back to grid tiling for {page_image_path.name}")

    step = max(256, tile_size - overlap)
    tiles = []

    y = 0
    row = 0
    while y < height:
        x = 0
        col = 0
        y2 = min(y + tile_size, height)
        while x < width:
            x2 = min(x + tile_size, width)
            tile = image.crop((x, y, x2, y2))
            tiles.append({
                "image": tile,
                "bbox": (x, y, x2, y2),
                "row": row,
                "col": col,
            })
            if x2 >= width:
                break
            x += step
            col += 1
        if y2 >= height:
            break
        y += step
        row += 1

    return tiles


def _extract_json_object(raw_text):
    """Extract a JSON object from raw model output."""
    if not raw_text:
        return None

    text = raw_text.strip()
    try:
        return json.loads(text)
    except Exception:
        pass

    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return None

    snippet = text[start:end + 1]
    try:
        return json.loads(snippet)
    except Exception:
        return None


def _validate_tile_bounds(tile_bounds, width, height, min_tiles):
    """Validate, clamp, and normalize tile bounds returned by the model."""
    if not tile_bounds or not isinstance(tile_bounds, list):
        return None

    valid = []
    for tile in tile_bounds:
        if not isinstance(tile, dict):
            continue
        try:
            x1 = int(tile.get("x1"))
            y1 = int(tile.get("y1"))
            x2 = int(tile.get("x2"))
            y2 = int(tile.get("y2"))
        except Exception:
            continue

        x1 = max(0, min(x1, width - 1))
        y1 = max(0, min(y1, height - 1))
        x2 = max(1, min(x2, width))
        y2 = max(1, min(y2, height))

        if x2 <= x1 or y2 <= y1:
            continue
        if (x2 - x1) < 64 or (y2 - y1) < 64:
            continue

        valid.append((x1, y1, x2, y2))

    # Deduplicate while preserving order.
    deduped = []
    seen = set()
    for bbox in valid:
        if bbox in seen:
            continue
        seen.add(bbox)
        deduped.append(bbox)

    if len(deduped) < min_tiles:
        return None

    return deduped


def _get_vision_guided_tile_bounds(page_image_path, width, height, min_tiles=6):
    """Ask the vision model for a JSON tile plan from the full page image."""
    try:
        import ollama
    except Exception as e:
        logging.warning(f"Ollama client not available for tile planning on {page_image_path.name}: {e}")
        return None

    vision_host = os.getenv("AI_VISION_HOST", "http://localhost:11434")
    candidate_models = [
        os.getenv("AI_VISION_MODEL"),
        "qwen3-vl:4b-instruct",
    ]
    candidate_models = [model for model in candidate_models if model]
    if not candidate_models:
        return None

    image_bytes = page_image_path.read_bytes()
    prompt = (
        "You are planning crop tiles for a dense engineering schematic page. "
        f"Image size is width={width}, height={height}. "
        f"Return a JSON object only with key 'tiles' containing at least {min_tiles} tiles. "
        "Each tile must include integer x1,y1,x2,y2 in pixel coordinates. "
        "Tiles should cover all major schematic regions and overlap neighboring tiles where needed for context. "
        "Ignore page border/frame lines and avoid border-only tiles. "
        "Prefer extra overlap around wiring intersections, connector blocks, and small text areas. "
        "Do not include markdown, prose, or comments. JSON only. "
        "Schema: {\"tiles\":[{\"x1\":0,\"y1\":0,\"x2\":1000,\"y2\":1000}]}"
    )

    last_error = None
    for model in candidate_models:
        try:
            client = ollama.Client(host=vision_host)
            response = client.chat(
                model=model,
                messages=[
                    {
                        "role": "user",
                        "content": prompt,
                        "images": [image_bytes],
                    }
                ],
                format="json",
            )

            content = None
            if hasattr(response, "message") and hasattr(response.message, "content"):
                content = response.message.content
            elif isinstance(response, dict):
                content = response.get("message", {}).get("content")

            parsed = _extract_json_object(content)
            if not parsed:
                continue

            plan = parsed.get("tiles") if isinstance(parsed, dict) else None
            bounds = _validate_tile_bounds(plan, width, height, min_tiles)
            if bounds:
                return bounds
        except Exception as e:
            last_error = e
            logging.warning(f"Vision tile plan model {model} failed for {page_image_path.name}: {e}")

    if last_error:
        logging.warning(f"No valid vision tile plan for {page_image_path.name}: {last_error}")
    return None


def describe_page_tile_with_vision(tile_image, tile_label, page_name):
    """Run a single tile through the vision model and return a strict schematic inventory."""
    try:
        import ollama
        import io

        vision_host = os.getenv("AI_VISION_HOST", "http://localhost:11434")
        candidate_models = [
            os.getenv("AI_VISION_MODEL"),
            "minicpm-v4.6",
        ]
        candidate_models = [model for model in candidate_models if model]

        buffer = io.BytesIO()
        tile_image.save(buffer, format="PNG")
        image_bytes = buffer.getvalue()

        prompt = (
            f"Page: {page_name}\n"
            f"Tile: {tile_label}\n\n"
            "Analyze this schematic tile only. Produce a strict inventory of what is visible in this tile. "
            "Transcribe all visible text exactly when possible. Preserve line breaks only where needed for readability. "
            "If multiple terminals, pins, connectors, or wire IDs are visible, list each one explicitly. "
            "Emphasize symbols such as switches, relays, fuses, sensors, terminals, connectors, PLC/I/O blocks, and wire junctions. "
            "Ignore page border/frame lines and title-block border graphics unless they contain actual signal labels. "
            "Do not guess hidden connections. If the element is partial or unreadable, say so. "
            "Return Markdown with these sections exactly: Visible Elements, Transcribed Labels, Inputs, Outputs, "
            "Wiring Configuration, Symbols and Components, Unclear or Illegible Text, and Notes."
        )

        last_error = None
        for model in candidate_models:
            try:
                client = ollama.Client(host=vision_host)
                response = client.chat(
                    model=model,
                    messages=[
                        {
                            "role": "user",
                            "content": prompt,
                            "images": [image_bytes],
                        }
                    ],
                )

                content = None
                if hasattr(response, "message") and hasattr(response.message, "content"):
                    content = response.message.content
                elif isinstance(response, dict):
                    content = response.get("message", {}).get("content")

                if content and content.strip():
                    return content.strip(), model
            except Exception as e:
                last_error = e
                logging.warning(f"Vision model {model} failed for {page_name} {tile_label}: {e}")

        if last_error:
            logging.warning(f"No tile description generated for {page_name} {tile_label}: {last_error}")
        return None, None

    except Exception as e:
        logging.warning(f"Could not run tile vision for {page_name} {tile_label}: {e}")
        return None, None


def merge_tile_descriptions(tile_records, page_name):
    """Merge tile inventories into one page-level schematic report."""
    if not tile_records:
        return None

    merged = []
    merged.append(f"# {page_name} Schematic Inventory")
    merged.append("")
    merged.append("## Method")
    merged.append("This page was split into overlapping tiles. Each tile was analyzed independently by a vision model, then merged into this report.")
    merged.append("")

    seen_lines = set()
    for record in tile_records:
        merged.append(f"## Tile {record['tile_label']}")
        merged.append(f"Tile bounds: {record['bbox']}")
        merged.append(f"Vision model: {record.get('model') or 'unknown'}")
        merged.append("")

        text = record.get("description") or ""
        if not text.strip():
            continue

        for line in text.splitlines():
            normalized = line.strip()
            if not normalized:
                merged.append("")
                continue
            if normalized.startswith("#"):
                continue
            if normalized in seen_lines:
                continue
            seen_lines.add(normalized)
            merged.append(normalized)

        merged.append("")

    return "\n".join(merged).rstrip() + "\n"


def write_page_schematic_descriptions(page_image_dir):
    """Tile every page image, run each tile through a vision model, and save a merged schematic report."""
    if not page_image_dir.exists():
        return 0

    page_images = sorted(page_image_dir.glob("page_*.png"))
    if not page_images:
        return 0

    saved_count = 0
    for page_image_path in page_images:
        tiles = _tile_page_image(page_image_path)
        if not tiles:
            continue

        tile_records = []
        for idx, tile in enumerate(tiles, start=1):
            x1, y1, x2, y2 = tile["bbox"]
            tile_label = f"r{tile['row'] + 1}_c{tile['col'] + 1}_t{idx:02d}"
            description, model_name = describe_page_tile_with_vision(
                tile["image"],
                tile_label,
                page_image_path.name,
            )
            if not description:
                continue
            tile_records.append({
                "tile_label": tile_label,
                "bbox": f"({x1}, {y1}, {x2}, {y2})",
                "description": description,
                "model": model_name,
            })

        merged = merge_tile_descriptions(tile_records, page_image_path.stem)
        if not merged:
            continue

        output_path = page_image_path.with_suffix(".schematic.md")
        try:
            output_path.write_text(merged, encoding="utf-8")
            saved_count += 1
        except Exception as e:
            logging.warning(f"Could not save schematic description for {page_image_path.name}: {e}")

    if saved_count > 0:
        logging.info(f"Saved {saved_count} schematic descriptions to {page_image_dir}")

    return saved_count

def get_file_extension(file_path):

    path = Path(file_path)
    
    # Primary method: use file extension from filename
    ext_from_name = path.suffix.lower()
    
    # Fallback method: use MIME type detection
    mime_type, _ = mimetypes.guess_type(str(path))
    ext_from_mime = None
    
    if mime_type:
        # Common MIME type to extension mappings
        mime_to_ext = {
            'application/pdf': '.pdf',
            'application/msword': '.doc',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
            'application/vnd.ms-excel': '.xls',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
            'application/vnd.ms-powerpoint': '.ppt',
            'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
            'text/csv': '.csv',
            'image/jpeg': '.jpg',
            'image/png': '.png',
            'image/gif': '.gif',
            'image/bmp': '.bmp',
            'image/webp': '.webp'
        }
        ext_from_mime = mime_to_ext.get(mime_type)
    
    # Use filename extension if available and valid, otherwise use MIME type extension
    if ext_from_name and ext_from_name in ALL_SUPPORTED_EXTENSIONS:
        return ext_from_name
    elif ext_from_mime and ext_from_mime in ALL_SUPPORTED_EXTENSIONS:
        logging.info(f"Using MIME type detection for {file_path}: {ext_from_mime}")
        return ext_from_mime
    else:
        # Return the filename extension even if not in our supported list
        # This allows the calling code to handle unsupported files appropriately
        return ext_from_name if ext_from_name else ''

def is_supported_file(file_path):

    """Check if the file type is supported for processing. Currently skips PDF files."""
    ext = get_file_extension(file_path)
    # Skip PDF files for now
    # if ext == '.pdf':
    #     return False
    return ext in ALL_SUPPORTED_EXTENSIONS

def process_pdf(pdf_path, doc_output_dir):
    pdf_file = Path(pdf_path)
    if not pdf_file.is_file():
        logging.warning(f"PDF file not found, skipping: {pdf_path}")
        return

    start_time = time.time()
    conv_res = doc_converter.convert(str(pdf_file))
    doc_filename = doc_output_dir.name
        
    # Ensure output directory exists
    doc_output_dir.mkdir(parents=True, exist_ok=True)

    # Export images to the output directory
    image_dir = doc_output_dir / "images"
    image_dir.mkdir(exist_ok=True)
    page_image_dir = doc_output_dir / "page_images"
    page_image_dir.mkdir(exist_ok=True)
    save_page_images(conv_res, page_image_dir)
    write_page_schematic_descriptions(page_image_dir)
    
    # Save images from the document
    image_count = 0
    for element, _ in conv_res.document.iterate_items():
        if isinstance(element, PictureItem):
            # Get the picture and save it
            picture = element.get_image(conv_res.document)
            if picture:
                # Sanitize the filename by replacing invalid characters
                safe_ref = element.self_ref.replace('#/', '').replace('/', '_').replace('\\', '_').replace(':', '_')
                image_filename = image_dir / f"{safe_ref}.png"
                
                # Ensure parent directory exists
                image_filename.parent.mkdir(parents=True, exist_ok=True)
                
                picture.save(image_filename)
                image_count += 1
                logging.debug(f"Saved image: {image_filename}")
    
    if image_count > 0:
        logging.info(f"Saved {image_count} images to {image_dir}")
    
    # Export markdown and ensure placeholders are replaced with saved image references.
    content_md = conv_res.document.export_to_markdown(image_mode=ImageRefMode.REFERENCED)
    content_md = inject_image_references(content_md, collect_saved_images(image_dir))
    
    # Save markdown
    md_filename = doc_output_dir / f"{doc_filename}.md"
    with md_filename.open("w", encoding="utf-8") as fp:
        fp.write(content_md)

    elapsed = time.time() - start_time
    logging.info(f"Converted {pdf_path} in {elapsed:.2f} seconds. Saved markdown to {md_filename}")

    return conv_res


def process_non_pdf(input_path, doc_output_dir):
    """Process a non-PDF file directly with docling."""
    input_file = Path(input_path)
    if not input_file.is_file():
        logging.warning(f"Input file not found, skipping: {input_path}")
        return

    start_time = time.time()
    doc_filename = doc_output_dir.name

    # Ensure output directory exists
    doc_output_dir.mkdir(parents=True, exist_ok=True)
    
    # Save document as JSON using model_dump_json (native docling serialization)
    json_filename = doc_output_dir / f"{doc_filename}_DoclingDocument.json"

    # Try to load existing JSON
    conv_res = None

    if json_filename.exists():
        try:
            logging.info(f"Attempting to load existing DoclingDocument from JSON...")
            json_content = json_filename.read_text(encoding="utf-8")
            
            # Load the document using the new docling 2.57.0 API
            doc = DoclingDocument.model_validate_json(json_content)
            
            # Create a minimal wrapper to mimic ConversionResult
            # This avoids needing the full ConversionResult object
            class DocumentWrapper:
                def __init__(self, document):
                    self.document = document
            
            conv_res = DocumentWrapper(doc)
            logging.info(f"✓ Successfully loaded DoclingDocument from JSON")

            #check if the markdown file also exists
            md_filename = doc_output_dir / f"{doc_filename}.md"
            if not md_filename.exists():
                logging.info(f"Markdown file not found, will regenerate from loaded document.")
                
                # Export images
                image_dir = doc_output_dir / "images"
                image_dir.mkdir(exist_ok=True)
                page_image_dir = doc_output_dir / "page_images"
                page_image_dir.mkdir(exist_ok=True)
                save_page_images(conv_res, page_image_dir)
                write_page_schematic_descriptions(page_image_dir)
                for element, _ in conv_res.document.iterate_items():
                    if isinstance(element, PictureItem):
                        picture = element.get_image(conv_res.document)
                        if picture:
                            # Sanitize the filename
                            safe_ref = element.self_ref.replace('#/', '').replace('/', '_').replace('\\', '_').replace(':', '_')
                            image_filename = image_dir / f"{safe_ref}.png"
                            image_filename.parent.mkdir(parents=True, exist_ok=True)
                            picture.save(image_filename)
                
                content_md = conv_res.document.export_to_markdown(image_mode=ImageRefMode.REFERENCED)
                content_md = inject_image_references(content_md, collect_saved_images(image_dir))
                with md_filename.open("w", encoding="utf-8") as fp:
                    fp.write(content_md)
                logging.info(f"Saved enhanced markdown to {md_filename}")

        except Exception as e:
            logging.warning(f"✗ Could not load JSON (likely old version): {type(e).__name__}")
            logging.info(f"Will reconvert document with docling 2.57.0...")
            conv_res = None
    
    # If loading failed or file doesn't exist, convert the document
    if conv_res is None:
        conv_res = doc_converter.convert(str(input_file))
        logging.info(f"✓ Converted {input_file.name} to DoclingDocument")

        # Save the document as JSON using model_dump_json
        try:
            with json_filename.open("w", encoding="utf-8") as fp:
                fp.write(conv_res.document.model_dump_json())
            logging.info(f"Saved document JSON to {json_filename}")
        except Exception as e:
            logging.warning(f"Could not save JSON file: {e}")

        # Export images to the output directory
        image_dir = doc_output_dir / "images"
        image_dir.mkdir(exist_ok=True)
        page_image_dir = doc_output_dir / "page_images"
        page_image_dir.mkdir(exist_ok=True)
        save_page_images(conv_res, page_image_dir)
        write_page_schematic_descriptions(page_image_dir)
        
        # Save images from the document
        image_count = 0
        for element, _ in conv_res.document.iterate_items():
            if isinstance(element, PictureItem):
                picture = element.get_image(conv_res.document)
                if picture:
                    # Sanitize the filename by replacing invalid characters
                    safe_ref = element.self_ref.replace('#/', '').replace('/', '_').replace('\\', '_').replace(':', '_')
                    image_filename = image_dir / f"{safe_ref}.png"
                    
                    # Ensure parent directory exists
                    image_filename.parent.mkdir(parents=True, exist_ok=True)
                    
                    picture.save(image_filename)
                    image_count += 1
                    logging.debug(f"Saved image: {image_filename}")
        
        if image_count > 0:
            logging.info(f"Saved {image_count} images to {image_dir}")
        
        # Export markdown and ensure placeholders are replaced with saved image references.
        content_md = conv_res.document.export_to_markdown(image_mode=ImageRefMode.REFERENCED)
        content_md = inject_image_references(content_md, collect_saved_images(image_dir))
        
        # Save markdown
        md_filename = doc_output_dir / f"{doc_filename}.md"
        with md_filename.open("w", encoding="utf-8") as fp:
            fp.write(content_md)
        logging.info(f"Saved enhanced markdown to {md_filename}")

    elapsed = time.time() - start_time
    logging.info(f"Processed {input_path} in {elapsed:.2f} seconds")

    return conv_res


def process_file(input_path, output_root):
 
    input_file = Path(input_path)
    
    # Skip temporary Office files (start with ~$)
    if input_file.name.startswith('~$'):
        logging.info(f"Skipping temporary Office file: {input_file}")
        return False
    
    ext = get_file_extension(input_file)
    
    if ext in PDF_EXTENSIONS:

        convert_res = process_pdf(str(input_file), output_root / input_file.stem)

        return convert_res


    elif is_supported_file(input_file):

        convert_res = process_non_pdf(str(input_file), output_root / input_file.stem)
        
        return convert_res
        
def main():

    # Create log file in the output directory
    output_root = Path(r"C:\Users\rperry\source\repos\Thesis Experiments\Testing\Test_Documents_Markdown")
    output_root.mkdir(parents=True, exist_ok=True)
    input_root = Path(r"C:\Users\rperry\source\repos\Thesis Experiments\Testing\Test_Documents")
    Testing_Root = Path(r"C:\Users\rperry\source\repos\Thesis Experiments\Testing")

    # Validate input directory
    if not input_root.exists():
        logging.error(f"Input directory does not exist: {input_root}")
        return

    print(f"Starting batch conversion...")
    print(f"Input directory: {input_root}")
    print(f"Output directory: {output_root}")
    print(f"PDF Pipeline options:")
    print(f"  - do_picture_description: {pdf_pipeline_options.do_picture_description}")
    print(f"  - generate_table_images: {pdf_pipeline_options.generate_table_images}")
    print(f"  - generate_picture_images: {pdf_pipeline_options.generate_picture_images}")
    print(f"  - images_scale: {pdf_pipeline_options.images_scale}")

    processed_count = 0
    error_count = 0
    pdf_converted = []
    pdf_not_converted = []

    # Process all files recursively
    for file_path in input_root.rglob('*'):
        if file_path.is_file() and not file_path.name.startswith('~$'):
            print(f"\n{'='*60}")
            print(f"Processing file: {file_path}")
            print(f"{'='*60}")

            # Get the file path folder structure relative to input_root
            relative_path = file_path.relative_to(input_root).parent
            output_subdir = output_root / relative_path
            
            # try:
            converted = process_file(file_path, output_subdir)
            processed_count += 1
            
    example_json = converted.document.model_dump_json()

    print(f"\nExample of converted document JSON structure saved to:")
    example_json_path = Testing_Root / "example_converted_document.json"
    with example_json_path.open("w", encoding="utf-8") as fp:
        fp.write(example_json)
    print(f"\nExample of converted document JSON structure saved to: {example_json_path}")

    print(f"\nBatch conversion completed.")
    print(f"Files processed: {processed_count}")
    print(f"Errors encountered: {error_count}")

    # Summary of PDF conversions
    print("\nPDF conversion summary:")
    logging.info("Processing complete.")

if __name__ == "__main__":
    main()