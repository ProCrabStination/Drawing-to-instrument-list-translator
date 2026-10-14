"""Test different OCR engines and VLM models with fixed document processing settings."""

import argparse
import json
import logging
import sys
import tempfile
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional, Dict, Any, Literal

try:
    import fitz
except ImportError:
    fitz = None

try:
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.base_models import InputFormat
    from docling.pipeline.vlm_pipeline import VlmPipeline
    from docling.datamodel.pipeline_options import (
        VlmPipelineOptions,
        VlmConvertOptions,
        PdfPipelineOptions,
        TesseractOcrOptions,
        EasyOcrOptions,
        RapidOcrOptions,
        OcrAutoOptions,
        OcrMacOptions,
    )
    from docling_core.types.doc.base import ImageRefMode
    DOCLING_AVAILABLE = True
except ImportError as e:
    DOCLING_AVAILABLE = False
    _DOCLING_IMPORT_ERROR = str(e)


@dataclass
class ConfigResult:
    """Result of running a single OCR/VLM configuration."""
    config_name: str
    test_type: Literal["ocr", "vlm"]
    model_name: str
    page_number: int
    success: bool
    duration_seconds: float
    output_chars: int
    output_file: Optional[str] = None
    error_message: Optional[str] = None
    instrument_ids_found: int = 0
    markdown_content: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return asdict(self)


class DoclingConfigTester:
    """Test harness for Docling OCR engines and VLM models."""

    # Fixed settings for all tests
    FIXED_SETTINGS = {
        "do_table_structure": True,
        "do_picture_description": True,
        "images_scale": 10.0,  # Maximum scale for maximum OCR detail on small instrument tags
    }

    # OCR engines - available in Docling 2.97.0
    # Note: SuryaOCR not available in current version
    ALL_OCR_ENGINES = [
        "auto",            # Auto-select best available
        "tesseract",       # Universal, widely available
        "easycor",         # PyTorch-based, good for complex layouts
        "rapidocr",        # ONNX/OpenVINO, fast
        "macos_native",    # macOS native (platform-specific)
    ]

    # VLM models organized by tier (recommended subset: 9 of 13)
    # Excludes: phi4 (licensing), gemma_27b (27B extreme resource),
    #           deepseek_ocr (API-only), unlimited_ocr (cloud dependency)
    VLM_TIER_1_LIGHTWEIGHT = [
        "smoldocling",            # SmolDocling-256M (~30s, fastest)
        "granite_docling",        # Granite-Docling-256M (~40s)
    ]

    VLM_TIER_2_SMALL = [
        # "granite_vision",         # Granite-Vision-3.3-2B (~2-3 mins)
        # "qwen",                   # Qwen2.5-VL-3B (~2-3 mins)
        # "nanonets_ocr2",          # Nanonets-OCR2-3B (~2-3 mins)
    ]

    VLM_TIER_3_MEDIUM = [
        # "got_ocr",                # GOT-OCR-2.0 (~3-5 mins)
        # "pixtral",                # Pixtral-12B (~8-12 mins)
    ]

    VLM_TIER_4_LARGE = [
        # "gemma_12b",              # Gemma-3-12B (12B params, MLX-only, ~8-12 mins)
        # "dolphin",                # Dolphin (~5-10 mins)
    ]

    # Predefined test mode configurations
    TEST_MODES = {
        "quick-test": {
            "ocr_engines": ["auto", "tesseract"],
            "vlm_models": ["granite_vision"],
            "description": "Quick validation (5-10 mins)",
        },
        "ocr-only": {
            "ocr_engines": ALL_OCR_ENGINES,
            "vlm_models": [],
            "description": "OCR benchmarking only (15-25 mins)",
        },
        "vlm-only": {
            "ocr_engines": [],
            "vlm_models": VLM_TIER_1_LIGHTWEIGHT + VLM_TIER_2_SMALL + VLM_TIER_3_MEDIUM + VLM_TIER_4_LARGE,
            "description": "VLM benchmarking only (60-120 mins)",
        },
        "production": {
            "ocr_engines": ALL_OCR_ENGINES,
            "vlm_models": VLM_TIER_1_LIGHTWEIGHT + VLM_TIER_2_SMALL + VLM_TIER_3_MEDIUM,
            "description": "Balanced benchmark (40-55 mins)",
        },
        "full-benchmark": {
            "ocr_engines": ALL_OCR_ENGINES,
            "vlm_models": VLM_TIER_1_LIGHTWEIGHT + VLM_TIER_2_SMALL + VLM_TIER_3_MEDIUM + VLM_TIER_4_LARGE,
            "description": "Complete suite (85-140 mins)",
        },
    }

    def __init__(self, output_dir: Optional[Path] = None, verbose: bool = False):
        self.output_dir = output_dir or Path.cwd() / "test_results"
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.verbose = verbose
        self.results: List[ConfigResult] = []
        self._setup_logging()

    def _setup_logging(self) -> None:
        """Configure logging."""
        level = logging.DEBUG if self.verbose else logging.INFO
        logging.basicConfig(
            level=level,
            format="[%(levelname)s] %(name)s: %(message)s"
        )
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _save_image_only_page(doc_pdf, page_number: int, out_path: Path, dpi: int = 300) -> None:
        """Rasterize one page into an image-only PDF so no embedded text layer reaches the converter."""
        source_page = doc_pdf[page_number - 1]
        pixmap = source_page.get_pixmap(dpi=dpi)
        output_pdf = fitz.open()
        new_page = output_pdf.new_page(width=source_page.rect.width, height=source_page.rect.height)
        new_page.insert_image(new_page.rect, pixmap=pixmap)
        output_pdf.save(str(out_path))
        output_pdf.close()

    def _extract_with_vlm_model(
        self,
        pdf_path: Path,
        page_number: int,
        vlm_model: str,
    ) -> Optional[str]:
        """Extract markdown from a single page using specified VLM model."""
        if not DOCLING_AVAILABLE:
            self.logger.error("Docling not available")
            return None

        if fitz is None:
            self.logger.error("PyMuPDF required but not installed")
            return None

        try:
            doc_pdf = fitz.open(str(pdf_path))
            if page_number < 1 or page_number > len(doc_pdf):
                self.logger.error(f"Page {page_number} out of range (1-{len(doc_pdf)})")
                doc_pdf.close()
                return None

            temp_pdf_path = Path(tempfile.mktemp(suffix=".pdf"))
            try:
                self._save_image_only_page(doc_pdf, page_number, temp_pdf_path)
                doc_pdf.close()

                # Use specified VLM model preset
                self.logger.info(f"Using VLM model preset: {vlm_model}")
                try:
                    vlm_convert_opts = VlmConvertOptions.from_preset(vlm_model)
                except Exception as e:
                    self.logger.warning(f"Failed to load VLM preset {vlm_model}: {e}")
                    self.logger.info("Falling back to default granite_docling preset")
                    vlm_convert_opts = VlmConvertOptions.from_preset("granite_docling")

                vlm_pipeline_options = VlmPipelineOptions(
                    vlm_options=vlm_convert_opts,
                    do_table_structure=self.FIXED_SETTINGS["do_table_structure"],
                    do_picture_description=self.FIXED_SETTINGS["do_picture_description"],
                    images_scale=self.FIXED_SETTINGS["images_scale"],
                )

                converter = DocumentConverter(
                    format_options={
                        InputFormat.PDF: PdfFormatOption(
                            pipeline_cls=VlmPipeline,
                            pipeline_options=vlm_pipeline_options,
                        ),
                    }
                )

                conv_result = converter.convert(str(temp_pdf_path))
                doc = conv_result.document
                markdown = doc.export_to_markdown(image_mode=ImageRefMode.REFERENCED)
                return markdown.strip() if markdown else None
            finally:
                if temp_pdf_path.exists():
                    temp_pdf_path.unlink()

        except Exception as exc:
            self.logger.error(f"Error with VLM {vlm_model} on page {page_number}: {exc}")
            if self.verbose:
                import traceback
                self.logger.debug(traceback.format_exc())
            return None

    def _extract_with_ocr_engine(
        self,
        pdf_path: Path,
        page_number: int,
        ocr_engine: str,
    ) -> Optional[str]:
        """Extract markdown from a single page using specified OCR engine via PdfPipeline."""
        if not DOCLING_AVAILABLE:
            self.logger.error("Docling not available")
            return None

        if fitz is None:
            self.logger.error("PyMuPDF required but not installed")
            return None

        try:
            doc_pdf = fitz.open(str(pdf_path))
            if page_number < 1 or page_number > len(doc_pdf):
                self.logger.error(f"Page {page_number} out of range (1-{len(doc_pdf)})")
                doc_pdf.close()
                return None

            temp_pdf_path = Path(tempfile.mktemp(suffix=".pdf"))
            try:
                self._save_image_only_page(doc_pdf, page_number, temp_pdf_path)
                doc_pdf.close()

                # Create OCR options for the specified engine
                self.logger.info(f"Using OCR engine: {ocr_engine}")
                ocr_options = self._create_ocr_options(ocr_engine)

                # Use standard PDF pipeline with specified OCR engine
                pdf_pipeline_options = PdfPipelineOptions(
                    ocr_options=ocr_options,
                    do_table_structure=self.FIXED_SETTINGS["do_table_structure"],
                    do_picture_description=self.FIXED_SETTINGS["do_picture_description"],
                    images_scale=self.FIXED_SETTINGS["images_scale"],
                )

                self.logger.info(f"Converting with {ocr_engine} OCR via PdfPipeline")
                converter = DocumentConverter(
                    format_options={
                        InputFormat.PDF: PdfFormatOption(
                            pipeline_options=pdf_pipeline_options,
                        ),
                    }
                )

                conv_result = converter.convert(str(temp_pdf_path))
                doc = conv_result.document
                markdown = doc.export_to_markdown(image_mode=ImageRefMode.REFERENCED)
                return markdown.strip() if markdown else None
            finally:
                if temp_pdf_path.exists():
                    temp_pdf_path.unlink()

        except Exception as exc:
            self.logger.error(f"Error with OCR {ocr_engine} on page {page_number}: {exc}")
            if self.verbose:
                import traceback
                self.logger.debug(traceback.format_exc())
            return None

    def _create_ocr_options(self, ocr_engine: str):
        """Create OCR options for the specified engine."""
        import platform

        if ocr_engine == "tesseract":
            return TesseractOcrOptions(
                lang=["eng"],
                force_full_page_ocr=True,
                bitmap_area_threshold=0.001,
            )
        elif ocr_engine == "easycor":
            return EasyOcrOptions(
                force_full_page_ocr=True,
                bitmap_area_threshold=0.001,
            )
        
        elif ocr_engine == "rapidocr":
            return RapidOcrOptions(
                lang=["eng"],
                force_full_page_ocr=True,
                bitmap_area_threshold=0.001,
            )
        # elif ocr_engine == "macos_native":
        #     # macOS native OCR only available on macOS
        #     if platform.system() != "Darwin":
        #         self.logger.warning("macOS native OCR only available on macOS, falling back to Tesseract")
        #         return TesseractOcrOptions(
        #         lang=["eng"],
        #         force_full_page_ocr=True,
        #         bitmap_area_threshold=0.001,
        #     )
        #     # Use macOS native OCR
        #     return OcrMacOptions()
        else:  # auto
            return OcrAutoOptions()

    def _count_instrument_ids(self, markdown: str) -> int:
        """Count detected instrument IDs in markdown (e.g., TI-101, LI-202, TI, PI, etc.)."""
        import re
        # Look for patterns like: TI-101, TI101, TI 101, PI-01, LI-2, but also single/paired letters like TI, PI, LI
        patterns = [
            r'\b[A-Z]{1,3}\s?-?\s?\d{1,4}\b',  # Full tags: TI-101, PI 01, etc.
            r'\b[A-Z]{1,2}\d{1,4}\b',          # Compact tags: TI101, PI5, etc.
            r'(?<![A-Z])[A-Z]{2}(?![A-Z])',    # Two-letter codes: TI, PI, LI, FI, etc. (isolated)
            r'(?<![A-Z])[A-Z]{1}(?![A-Z0-9])', # Single letter codes: T, P, L, F (isolated)
        ]
        all_ids = []
        for pattern in patterns:
            matches = re.findall(pattern, markdown, re.IGNORECASE)
            all_ids.extend(matches)
            if matches and len(matches) < 50:  # Only log if reasonable amount
                self.logger.debug(f"Found {len(matches)} with pattern '{pattern}': {set(matches)}")

        # Filter out common non-instrument abbreviations
        filtered = {id for id in all_ids if id.upper() not in ['DO', 'TO', 'GO', 'NO', 'SO', 'OR', 'OF', 'ON', 'BY', 'IF']}
        return len(filtered)

    def test_vlm_model(
        self,
        pdf_path: Path,
        page_number: int,
        vlm_model: str,
    ) -> ConfigResult:
        """Test a single VLM model and record results."""
        config_name = f"vlm_{vlm_model}"
        self.logger.info(f"Testing VLM model: {vlm_model}")
        self.logger.info(f"  Fixed settings: {self.FIXED_SETTINGS}")

        start_time = time.time()
        result = ConfigResult(
            config_name=config_name,
            test_type="vlm",
            model_name=vlm_model,
            page_number=page_number,
            success=False,
            duration_seconds=0.0,
            output_chars=0,
        )

        try:
            markdown = self._extract_with_vlm_model(pdf_path, page_number, vlm_model)

            duration = time.time() - start_time
            result.duration_seconds = duration

            if markdown:
                result.success = True
                result.output_chars = len(markdown)
                result.instrument_ids_found = self._count_instrument_ids(markdown)
                result.markdown_content = markdown

                output_file = self.output_dir / f"{config_name}_page{page_number}.md"
                output_file.write_text(markdown, encoding="utf-8")
                result.output_file = str(output_file)

                self.logger.info(f"  ✓ Success in {duration:.2f}s: {len(markdown)} chars, "
                                f"{result.instrument_ids_found} instrument IDs")
            else:
                result.error_message = "No markdown extracted"
                self.logger.warning(f"  ✗ Failed: {result.error_message}")

        except Exception as exc:
            duration = time.time() - start_time
            result.duration_seconds = duration
            result.error_message = str(exc)
            self.logger.error(f"  ✗ Failed: {result.error_message}")

        self.results.append(result)
        return result

    def test_ocr_engine(
        self,
        pdf_path: Path,
        page_number: int,
        ocr_engine: str,
    ) -> ConfigResult:
        """Test a single OCR engine and record results."""
        config_name = f"ocr_{ocr_engine}"
        self.logger.info(f"Testing OCR engine: {ocr_engine}")
        self.logger.info(f"  Fixed settings: {self.FIXED_SETTINGS}")

        start_time = time.time()
        result = ConfigResult(
            config_name=config_name,
            test_type="ocr",
            model_name=ocr_engine,
            page_number=page_number,
            success=False,
            duration_seconds=0.0,
            output_chars=0,
        )

        try:
            markdown = self._extract_with_ocr_engine(pdf_path, page_number, ocr_engine)

            duration = time.time() - start_time
            result.duration_seconds = duration

            if markdown:
                result.success = True
                result.output_chars = len(markdown)
                result.instrument_ids_found = self._count_instrument_ids(markdown)
                result.markdown_content = markdown

                output_file = self.output_dir / f"{config_name}_page{page_number}.md"
                output_file.write_text(markdown, encoding="utf-8")
                result.output_file = str(output_file)

                self.logger.info(f"  ✓ Success in {duration:.2f}s: {len(markdown)} chars, "
                                f"{result.instrument_ids_found} instrument IDs")
            else:
                result.error_message = "No markdown extracted"
                self.logger.warning(f"  ✗ Failed: {result.error_message}")

        except Exception as exc:
            duration = time.time() - start_time
            result.duration_seconds = duration
            result.error_message = str(exc)
            self.logger.error(f"  ✗ OCR engine '{ocr_engine}' failed: {result.error_message}")
            if self.verbose:
                import traceback
                self.logger.debug(traceback.format_exc())

        self.results.append(result)
        return result

    def generate_report(self) -> Dict[str, Any]:
        """Generate a summary report of all test results."""
        report = {
            "output_directory": str(self.output_dir),
            "fixed_settings": self.FIXED_SETTINGS,
            "total_tests": len(self.results),
            "successful": sum(1 for r in self.results if r.success),
            "failed": sum(1 for r in self.results if not r.success),
            "results": [r.to_dict() for r in self.results],
            "vlm_comparison": {},
            "ocr_comparison": {},
        }

        # Group by type and model
        vlm_results = [r for r in self.results if r.test_type == "vlm" and r.success]
        ocr_results = [r for r in self.results if r.test_type == "ocr" and r.success]

        # Collect all VLM models from tiers
        all_vlm_models = set()
        all_vlm_models.update(self.VLM_TIER_1_LIGHTWEIGHT)
        all_vlm_models.update(self.VLM_TIER_2_SMALL)
        all_vlm_models.update(self.VLM_TIER_3_MEDIUM)
        all_vlm_models.update(self.VLM_TIER_4_LARGE)

        # VLM comparison
        for model_name in all_vlm_models:
            matching = [r for r in vlm_results if r.model_name == model_name]
            if matching:
                durations = [r.duration_seconds for r in matching]
                chars = [r.output_chars for r in matching]
                ids = [r.instrument_ids_found for r in matching]
                report["vlm_comparison"][model_name] = {
                    "tests": len(matching),
                    "avg_duration": sum(durations) / len(durations),
                    "avg_output_chars": sum(chars) / len(chars),
                    "avg_instrument_ids": sum(ids) / len(ids),
                    "min_duration": min(durations),
                    "max_duration": max(durations),
                }

        # OCR comparison
        for engine_name in self.ALL_OCR_ENGINES:
            matching = [r for r in ocr_results if r.model_name == engine_name]
            if matching:
                durations = [r.duration_seconds for r in matching]
                chars = [r.output_chars for r in matching]
                ids = [r.instrument_ids_found for r in matching]
                report["ocr_comparison"][engine_name] = {
                    "tests": len(matching),
                    "avg_duration": sum(durations) / len(durations),
                    "avg_output_chars": sum(chars) / len(chars),
                    "avg_instrument_ids": sum(ids) / len(ids),
                    "min_duration": min(durations),
                    "max_duration": max(durations),
                }

        return report

    def save_report(self, output_file: Optional[Path] = None) -> Path:
        """Save the report as JSON."""
        if output_file is None:
            output_file = self.output_dir / "test_results.json"

        report = self.generate_report()
        output_file.write_text(json.dumps(report, indent=2), encoding="utf-8")
        self.logger.info(f"Report saved to: {output_file}")
        return output_file

    def print_summary(self) -> None:
        """Print a summary to console."""
        report = self.generate_report()

        print("\n" + "="*80)
        print("DOCLING OCR/VLM MODEL TEST SUMMARY")
        print("="*80)
        print(f"\nOutput directory: {report['output_directory']}")
        print(f"Total tests: {report['total_tests']}")
        print(f"Successful: {report['successful']}")
        print(f"Failed: {report['failed']}")
        print(f"\nFixed Settings:")
        for key, value in report['fixed_settings'].items():
            print(f"  {key}: {value}")

        if report["vlm_comparison"]:
            print("\n" + "-"*80)
            print("VLM Model Comparison (sorted by speed):")
            print("-"*80)
            for model, stats in sorted(report["vlm_comparison"].items(),
                                      key=lambda x: x[1]["avg_duration"]):
                print(f"\n{model}:")
                print(f"  Tests: {stats['tests']}")
                print(f"  Avg duration: {stats['avg_duration']:.2f}s (min: {stats['min_duration']:.2f}s, max: {stats['max_duration']:.2f}s)")
                print(f"  Avg output chars: {stats['avg_output_chars']:.0f}")
                print(f"  Avg instrument IDs found: {stats['avg_instrument_ids']:.0f}")

        if report["ocr_comparison"]:
            print("\n" + "-"*80)
            print("OCR Engine Comparison (sorted by speed):")
            print("-"*80)
            for engine, stats in sorted(report["ocr_comparison"].items(),
                                       key=lambda x: x[1]["avg_duration"]):
                print(f"\n{engine}:")
                print(f"  Tests: {stats['tests']}")
                print(f"  Avg duration: {stats['avg_duration']:.2f}s (min: {stats['min_duration']:.2f}s, max: {stats['max_duration']:.2f}s)")
                print(f"  Avg output chars: {stats['avg_output_chars']:.0f}")
                print(f"  Avg instrument IDs found: {stats['avg_instrument_ids']:.0f}")

        print("\n" + "="*80 + "\n")


def print_test_plan(mode: str = "production") -> None:
    """Print the test plan for the specified mode."""
    tester = DoclingConfigTester()
    mode_config = tester.TEST_MODES.get(mode, tester.TEST_MODES["production"])

    print("\n" + "="*80)
    print(f"DOCLING OCR/VLM BENCHMARK - {mode.upper()} MODE")
    print("="*80)
    print(f"\nMode: {mode}")
    print(f"Duration estimate: {mode_config['description']}")
    print("\nFixed Document Settings (All Tests):")
    for key, value in tester.FIXED_SETTINGS.items():
        print(f"  • {key}: {value}")

    print("\n" + "-"*80)
    print("RECOMMENDED VLM MODELS (9 of 13)")
    print("-"*80)

    print("\nTIER 1 - LIGHTWEIGHT (fastest, ~30-40s each):")
    for i, model in enumerate(tester.VLM_TIER_1_LIGHTWEIGHT, 1):
        status = "✓ TESTING" if model in mode_config["vlm_models"] else "  Skipped"
        print(f"  {status}  {i}. {model}")

    print("\nTIER 2 - SMALL (~2-3 mins each):")
    for i, model in enumerate(tester.VLM_TIER_2_SMALL, len(tester.VLM_TIER_1_LIGHTWEIGHT) + 1):
        status = "✓ TESTING" if model in mode_config["vlm_models"] else "  Skipped"
        print(f"  {status}  {i}. {model}")

    print("\nTIER 3 - MEDIUM (~3-12 mins each):")
    for i, model in enumerate(tester.VLM_TIER_3_MEDIUM, len(tester.VLM_TIER_1_LIGHTWEIGHT) + len(tester.VLM_TIER_2_SMALL) + 1):
        status = "✓ TESTING" if model in mode_config["vlm_models"] else "  Skipped"
        print(f"  {status}  {i}. {model}")

    print("\nTIER 4 - LARGE (~5-12 mins each, resource-intensive):")
    for i, model in enumerate(tester.VLM_TIER_4_LARGE, len(tester.VLM_TIER_1_LIGHTWEIGHT) + len(tester.VLM_TIER_2_SMALL) + len(tester.VLM_TIER_3_MEDIUM) + 1):
        status = "✓ TESTING" if model in mode_config["vlm_models"] else "  Skipped"
        print(f"  {status}  {i}. {model}")

    print("\nEXCLUDED MODELS (not recommended for benchmarking):")
    print("  • phi4              - Proprietary licensing complexity")
    print("  • gemma_27b         - Extreme resource requirements (27B params)")
    print("  • deepseek_ocr      - API-only, requires cloud credentials")
    print("  • unlimited_ocr     - Cloud dependency, incompatible with offline")

    print("\n" + "-"*80)
    print("OCR ENGINES (5 engines in Docling 2.97.0)")
    print("-"*80)
    for i, engine in enumerate(tester.ALL_OCR_ENGINES, 1):
        status = "✓ TESTING" if engine in mode_config["ocr_engines"] else "  Skipped"
        desc = {
            "auto": "Auto-select best available",
            "tesseract": "Universal OCR (widely available)",
            "easycor": "EasyOCR (PyTorch-based)",
            "rapidocr": "RapidOCR (ONNX/OpenVINO, fast)",
            "macos_native": "macOS native (platform-specific)",
        }.get(engine, "")
        print(f"  {status}  {i}. {engine:15} - {desc}")
    print("\nNOTE:")
    print("  • SuryaOCR not available in Docling 2.97.0")

    print("\n" + "-"*80)
    print("EXECUTION SUMMARY")
    print("-"*80)
    total_tests = len(mode_config["vlm_models"]) + len(mode_config["ocr_engines"])
    print(f"  Total tests in {mode} mode: {total_tests}")
    print(f"    • VLM Models: {len(mode_config['vlm_models'])}")
    print(f"    • OCR Engines: {len(mode_config['ocr_engines'])}")
    print(f"  Metrics per test:")
    print(f"    • Execution time (seconds)")
    print(f"    • Output size (characters)")
    print(f"    • Instrument IDs detected")
    print(f"    • Success/failure status")
    print("\n" + "="*80 + "\n")


def main() -> int:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Comprehensive Docling OCR/VLM benchmark suite (9 VLM + 5 OCR engines)."
    )
    parser.add_argument(
        "--pdf",
        type=Path,
        help="PDF file to test (auto-detect if not specified)"
    )
    parser.add_argument(
        "--page",
        type=int,
        default=1,
        help="Page number to test (default: 1)"
    )
    parser.add_argument(
        "--mode",
        choices=["quick-test", "ocr-only", "vlm-only", "production", "full-benchmark"],
        default="production",
        help="Test mode (default: production - 5 OCR + 7 VLM models, ~45-60 mins)"
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Directory for output files (default: temp directory)"
    )
    parser.add_argument(
        "--plan",
        action="store_true",
        help="Show the test plan for the selected mode and exit"
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable verbose logging"
    )

    args = parser.parse_args()

    # Initialize tester
    tester = DoclingConfigTester(output_dir=args.output_dir, verbose=args.verbose)

    # Show plan if requested
    if args.plan:
        print_test_plan(args.mode)
        return 0

    # Find PDF if not specified
    pdf_path = args.pdf
    if not pdf_path:
        current_dir = Path(".")
        pdfs = list(current_dir.glob("*.pdf"))
        if not pdfs:
            print("Error: No PDF files found in current directory.")
            print("Usage: python test_docling_configurations.py --pdf <path/to/file.pdf>")
            print("       python test_docling_configurations.py --plan")
            return 1
        pdf_path = pdfs[0]
        print(f"Auto-detected PDF: {pdf_path}\n")

    if not pdf_path.exists():
        print(f"Error: PDF file not found: {pdf_path}")
        return 1

    # Get mode configuration
    mode_config = tester.TEST_MODES.get(args.mode, tester.TEST_MODES["production"])

    # Print the test plan
    print_test_plan(args.mode)
    print(f"Starting benchmark on: {pdf_path.name} (page {args.page})")
    print(f"Output directory: {tester.output_dir}\n")

    test_count = 0

    # Test OCR engines
    if mode_config["ocr_engines"]:
        print(f"{'='*80}")
        print(f"PHASE 1: Testing {len(mode_config['ocr_engines'])} OCR Engines")
        print(f"{'='*80}\n")
        for i, ocr_engine in enumerate(mode_config["ocr_engines"], 1):
            print(f"[{i}/{len(mode_config['ocr_engines'])}] {ocr_engine}...", flush=True)
            tester.test_ocr_engine(pdf_path, args.page, ocr_engine)
            test_count += 1

    # Test VLM models in tiers
    if mode_config["vlm_models"]:
        print(f"\n{'='*80}")
        print(f"PHASE 2: Testing {len(mode_config['vlm_models'])} VLM Models")
        print(f"{'='*80}\n")
        for i, vlm_model in enumerate(mode_config["vlm_models"], 1):
            print(f"[{i}/{len(mode_config['vlm_models'])}] {vlm_model}...", flush=True)
            tester.test_vlm_model(pdf_path, args.page, vlm_model)
            test_count += 1

    print(f"\n{'='*80}")
    print(f"✓ All {test_count} tests completed")
    print(f"{'='*80}")
    print(f"Output directory: {tester.output_dir}\n")

    # Generate and save report
    report_path = tester.save_report()
    tester.print_summary()

    print(f"📊 Full report saved to: {report_path}\n")

    return 0


if __name__ == "__main__":
    sys.exit(main())
