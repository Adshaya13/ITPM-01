from __future__ import annotations

import argparse
import base64
import csv
import io
import sys
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from playwright.sync_api import sync_playwright


BASE_URL = "https://www.pixelssuite.com"
DEFAULT_TIMEOUT_MS = 20000
DEFAULT_SLOW_MO_MS = 200
RESULTS_DIR = Path("results")
CSV_PATH = Path("execution_results.csv")

PNG_1X1_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB/6X9wYQAAAAASUVORK5CYII="
)

JPG_1X1_BASE64 = (
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wCEAAkGBxISEhUQEhIVFRUVFRUVFRUVFRUVFRUXFhUVFRUYHSgg"
    "GBolHRUVITEhJSkrLi4uFx8zODMtNygtLisBCgoKDg0OFxAQFy0dHR0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLf/AABEIAAEAAQMBIgACEQEDEQH/xAAbAAACAgMBAAAAAAAAAAAAAAAFBgMEAAIHCf/EADYQAAEDAgQDBgQEBQUAAAAAAAEAAgMEEQUSITEGE0FRYXEicYEykaGxwSNSwdHh8CMzU2KissLx/8QAGAEAAwEBAAAAAAAAAAAAAAAAAQIDBAX/xAAlEQACAgICAgICAgMAAAAAAAAAAQIRAyESMQRBIlETcYGR8P/aAAwDAQACEQMRAD8A+6iiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA//2Q=="
)

GIF_1X1_BASE64 = "R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw=="


ROUTES = {
    "image_to_pdf": "/image-to-pdf",
    "pdf_to_word": "/pdf-to-word",
    "word_to_pdf": "/word-to-pdf",
    "pdf_editor": "/pdf-editor",
    "resize_image": "/resize-image",
    "bulk_resize": "/bulk-resize",
    "image_enlarger": "/image-enlarger",
    "crop_jpg": "/crop-jpg",
    "crop_png": "/crop-png",
    "crop_webp": "/crop-webp",
    "compress_image": "/compress-image",
    "png_compressor": "/png-compressor",
    "gif_compressor": "/gif-compressor",
    "convert_image": "/convert-image",
    "convert_jpg": "/convert-to-jpg",
    "convert_png": "/convert-to-png",
    "convert_webp": "/convert-to-webp",
    "rotate_image": "/rotate-image",
    "flip_image": "/flip-image",
    "meme_generator": "/meme-generator",
    "color_picker": "/color-picker",
}


def configure_stdout() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run PixelsSuite Playwright tests.")
    parser.add_argument("--base-url", default=BASE_URL)
    parser.add_argument("--headless", action="store_true", default=False)
    parser.add_argument("--timeout-ms", type=int, default=DEFAULT_TIMEOUT_MS)
    parser.add_argument("--slow-mo-ms", type=int, default=DEFAULT_SLOW_MO_MS)
    parser.add_argument("--only", default="", help="Comma-separated TC_ID list to run.")
    return parser.parse_args()


def decode_base64_file(destination: Path, data: str) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if not destination.exists():
        destination.write_bytes(base64.b64decode(data))


def build_sample_pdf_bytes() -> bytes:
    content_stream = b"BT /F1 24 Tf 100 700 Td (Sample PDF) Tj ET\n"
    objects = [
        b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
        b"2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n",
        (
            b"3 0 obj\n"
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            b"/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n"
        ),
        (
            b"4 0 obj\n"
            + f"<< /Length {len(content_stream)} >>\n".encode("ascii")
            + b"stream\n"
            + content_stream
            + b"endstream\nendobj\n"
        ),
        b"5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n",
    ]

    buffer = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for obj in objects:
        offsets.append(len(buffer))
        buffer.extend(obj)

    xref_offset = len(buffer)
    buffer.extend(f"xref\n0 {len(offsets)}\n".encode("ascii"))
    buffer.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        buffer.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    buffer.extend(
        (
            "trailer\n"
            f"<< /Size {len(offsets)} /Root 1 0 R >>\n"
            "startxref\n"
            f"{xref_offset}\n"
            "%%EOF\n"
        ).encode("ascii")
    )
    return bytes(buffer)


def build_sample_docx_bytes() -> bytes:
        try:
                from docx import Document

                buffer = io.BytesIO()
                document = Document()
                document.add_paragraph("Sample DOCX")
                document.save(buffer)
                return buffer.getvalue()
        except Exception:
                document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>Sample DOCX</w:t>
            </w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
    </w:body>
</w:document>
"""

                content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

                rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

                document_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"""

                buffer = io.BytesIO()
                with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
                        archive.writestr("[Content_Types].xml", content_types_xml)
                        archive.writestr("_rels/.rels", rels_xml)
                        archive.writestr("word/document.xml", document_xml)
                        archive.writestr("word/_rels/document.xml.rels", document_rels_xml)
                return buffer.getvalue()


def ensure_fixture_files(workspace_root: Path) -> dict[str, Path]:
    fixtures = {
        "sample_png": workspace_root / "sample.png",
        "sample_jpg": workspace_root / "sample.jpg",
        "sample_gif": workspace_root / "sample.gif",
        "sample_pdf": workspace_root / "sample.pdf",
        "sample_docx": workspace_root / "sample.docx",
        "invalid_txt": workspace_root / "invalid.txt",
    }

    decode_base64_file(fixtures["sample_png"], PNG_1X1_BASE64)
    decode_base64_file(fixtures["sample_jpg"], JPG_1X1_BASE64)
    decode_base64_file(fixtures["sample_gif"], GIF_1X1_BASE64)

    if not fixtures["sample_pdf"].exists():
        fixtures["sample_pdf"].write_bytes(build_sample_pdf_bytes())

    fixtures["sample_docx"].write_bytes(build_sample_docx_bytes())

    if not fixtures["invalid_txt"].exists():
        fixtures["invalid_txt"].write_text("This is an invalid test file.\n", encoding="utf-8")

    return fixtures


def write_csv_header(csv_path: Path) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=["TC_ID", "Feature", "Status", "Screenshot"])
        writer.writeheader()


def append_csv_row(csv_path: Path, row: dict[str, str]) -> None:
    with csv_path.open("a", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=["TC_ID", "Feature", "Status", "Screenshot"])
        writer.writerow(row)


def normalize_path(path: str, base_url: str) -> str:
    if path.startswith("http://") or path.startswith("https://"):
        return path
    return f"{base_url.rstrip('/')}{path}"


def wait_for_condition(page, predicate: Callable[[], bool], timeout_ms: int, message: str) -> None:
    deadline = time.monotonic() + (timeout_ms / 1000)
    last_error: Exception | None = None

    while time.monotonic() < deadline:
        try:
            if predicate():
                return
        except Exception as exc:
            last_error = exc
        page.wait_for_timeout(250)

    if last_error:
        raise AssertionError(message) from last_error
    raise AssertionError(message)


def upload_files(page, file_paths: list[Path]) -> None:
    file_input = page.locator('input[type="file"]').first
    file_input.set_input_files([str(path) for path in file_paths])
    page.wait_for_timeout(500)


def button(page, name: str, exact: bool = False):
    return page.get_by_role("button", name=name, exact=exact)


def body_text(page) -> str:
    return page.locator("body").inner_text(timeout=5000)


def assert_text_contains(page, expected: str) -> None:
    text = body_text(page).lower()
    if expected.lower() not in text:
        raise AssertionError(f'Expected page text to contain "{expected}"')


def assert_button_disabled(page, name: str) -> None:
    locator = button(page, name)
    if locator.count() == 0:
        return
    if locator.is_enabled():
        raise AssertionError(f'Expected button "{name}" to be disabled')


def assert_button_enabled(page, name: str, timeout_ms: int) -> None:
    wait_for_condition(
        page,
        lambda: button(page, name).is_enabled(),
        timeout_ms,
        f'Expected button "{name}" to become enabled',
    )


def expect_download_on_click(page, name: str, timeout_ms: int) -> str:
    with page.expect_download(timeout=timeout_ms) as download_info:
        button(page, name).click()
    return download_info.value.suggested_filename


@dataclass(frozen=True)
class TestCase:
    tc_id: str
    feature: str
    name: str
    route: str
    runner: Callable[[object, dict[str, Path], int], None]


TEST_CASES: list[TestCase] = []


def register_case(tc_id: str, feature: str, name: str, route: str):
    def decorator(fn: Callable[[object, dict[str, Path], int], None]):
        TEST_CASES.append(TestCase(tc_id=tc_id, feature=feature, name=name, route=route, runner=fn))
        globals()[name] = fn
        return fn

    return decorator


@register_case("TC01", "Document conversion", "tc01_image_to_pdf_valid", ROUTES["image_to_pdf"])
def tc01_image_to_pdf_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"], fixtures["sample_jpg"]])
    assert_button_enabled(page, "Create PDF", timeout_ms)
    assert_text_contains(page, "sample.png")
    assert_text_contains(page, "sample.jpg")
    filename = expect_download_on_click(page, "Create PDF", timeout_ms)
    if not filename.lower().endswith(".pdf"):
        raise AssertionError("Image to PDF should download a PDF file")


@register_case("TC02", "Document conversion", "tc02_image_to_pdf_empty", ROUTES["image_to_pdf"])
def tc02_image_to_pdf_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Create PDF")


@register_case("TC03", "Document conversion", "tc03_pdf_to_word_valid", ROUTES["pdf_to_word"])
def tc03_pdf_to_word_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_pdf"]])
    assert_button_enabled(page, "Convert to Word", timeout_ms)
    button(page, "Convert to Word", exact=True).click()
    assert_text_contains(page, "Converting…")


@register_case("TC04", "Document conversion", "tc04_pdf_to_word_invalid", ROUTES["pdf_to_word"])
def tc04_pdf_to_word_invalid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["invalid_txt"]])
    assert_button_enabled(page, "Convert to Word", timeout_ms)
    try:
        expect_download_on_click(page, "Convert to Word", timeout_ms)
    except Exception:
        pass
    else:
        raise AssertionError("Invalid PDF input should not download a DOCX")
    assert_text_contains(page, "conversion failed")


@register_case("TC05", "Document conversion", "tc05_word_to_pdf_valid", ROUTES["word_to_pdf"])
def tc05_word_to_pdf_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_docx"]])
    assert_button_enabled(page, "Convert to PDF", timeout_ms)
    button(page, "Convert to PDF", exact=True).click()
    assert_text_contains(page, "Converting…")


@register_case("TC06", "Document conversion", "tc06_word_to_pdf_invalid", ROUTES["word_to_pdf"])
def tc06_word_to_pdf_invalid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["invalid_txt"]])
    assert_button_enabled(page, "Convert to PDF", timeout_ms)
    try:
        expect_download_on_click(page, "Convert to PDF", timeout_ms)
    except Exception:
        pass
    else:
        raise AssertionError("Invalid DOCX input should not download a PDF")
    assert_text_contains(page, "conversion failed")


@register_case("TC07", "PDF editing", "tc07_pdf_editor_open", ROUTES["pdf_editor"])
def tc07_pdf_editor_open(page, fixtures, timeout_ms):
    assert_text_contains(page, "PDF Editor")
    if page.locator('input[type="file"]').count() == 0:
        raise AssertionError("Expected PDF editor to expose a file input")
    assert_button_disabled(page, "Download")


@register_case("TC08", "PDF editing", "tc08_pdf_editor_valid", ROUTES["pdf_editor"])
def tc08_pdf_editor_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_pdf"]])
    assert_button_enabled(page, "Download", timeout_ms)
    filename = expect_download_on_click(page, "Download", timeout_ms)
    if not filename.lower().endswith(".pdf"):
        raise AssertionError("PDF editor should download a PDF file")


@register_case("TC09", "PDF editing", "tc09_pdf_editor_empty", ROUTES["pdf_editor"])
def tc09_pdf_editor_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Download")


@register_case("TC10", "Image resizing", "tc10_resize_image_valid", ROUTES["resize_image"])
def tc10_resize_image_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download PNG", timeout_ms)
    filename = expect_download_on_click(page, "Download PNG", timeout_ms)
    if not filename.lower().endswith(".png"):
        raise AssertionError("Resize Image should download a PNG file")


@register_case("TC11", "Image resizing", "tc11_resize_image_empty", ROUTES["resize_image"])
def tc11_resize_image_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Download PNG")


@register_case("TC12", "Image resizing", "tc12_bulk_resize_valid", ROUTES["bulk_resize"])
def tc12_bulk_resize_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"], fixtures["sample_jpg"]])
    assert_button_enabled(page, "Process & Download", timeout_ms)
    filename = expect_download_on_click(page, "Process & Download", timeout_ms)
    if not filename.lower().endswith(".png"):
        raise AssertionError("Bulk Resize should download PNG files")


@register_case("TC13", "Image resizing", "tc13_bulk_resize_empty", ROUTES["bulk_resize"])
def tc13_bulk_resize_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Process & Download")


@register_case("TC14", "Image resizing", "tc14_image_enlarger_valid", ROUTES["image_enlarger"])
def tc14_image_enlarger_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download PNG", timeout_ms)
    filename = expect_download_on_click(page, "Download PNG", timeout_ms)
    if not filename.lower().endswith(".png"):
        raise AssertionError("Image Enlarger should download a PNG file")


@register_case("TC15", "Image resizing", "tc15_image_enlarger_empty", ROUTES["image_enlarger"])
def tc15_image_enlarger_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Download PNG")


@register_case("TC16", "Cropping", "tc16_crop_jpg_valid", ROUTES["crop_jpg"])
def tc16_crop_jpg_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download", timeout_ms)
    filename = expect_download_on_click(page, "Download", timeout_ms)
    if not filename.lower().endswith((".jpg", ".jpeg")):
        raise AssertionError("Crop JPG should download a JPG file")


@register_case("TC17", "Cropping", "tc17_crop_jpg_empty", ROUTES["crop_jpg"])
def tc17_crop_jpg_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Download")


@register_case("TC18", "Cropping", "tc18_crop_png_valid", ROUTES["crop_png"])
def tc18_crop_png_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_jpg"]])
    assert_text_contains(page, "Crop PNG")
    assert_text_contains(page, "Select an image to crop.")
    assert_text_contains(page, "Clear")


@register_case("TC19", "Cropping", "tc19_crop_webp_valid", ROUTES["crop_webp"])
def tc19_crop_webp_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download", timeout_ms)
    filename = expect_download_on_click(page, "Download", timeout_ms)
    if not filename.lower().endswith(".webp"):
        raise AssertionError("Crop WebP should download a WebP file")


@register_case("TC20", "Compression", "tc20_compress_image_valid", ROUTES["compress_image"])
def tc20_compress_image_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download", timeout_ms)
    filename = expect_download_on_click(page, "Download", timeout_ms)
    if not any(filename.lower().endswith(ext) for ext in (".jpg", ".jpeg", ".webp")):
        raise AssertionError("Compress Image should download a compressed image")


@register_case("TC21", "Compression", "tc21_compress_image_empty", ROUTES["compress_image"])
def tc21_compress_image_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Download")


@register_case("TC22", "Compression", "tc22_png_compressor_valid", ROUTES["png_compressor"])
def tc22_png_compressor_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download PNG", timeout_ms)
    filename = expect_download_on_click(page, "Download PNG", timeout_ms)
    if not filename.lower().endswith(".png"):
        raise AssertionError("PNG Compressor should download a PNG file")


@register_case("TC23", "Compression", "tc23_gif_compressor_valid", ROUTES["gif_compressor"])
def tc23_gif_compressor_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_gif"]])
    wait_for_condition(
        page,
        lambda: button(page, "Compress", exact=True).count() > 0 and button(page, "Compress", exact=True).is_enabled(),
        timeout_ms,
        'Expected button "Compress" to become enabled',
    )
    button(page, "Compress", exact=True).click()
    assert_button_enabled(page, "Download GIF", timeout_ms)
    filename = expect_download_on_click(page, "Download GIF", timeout_ms)
    if not filename.lower().endswith(".gif"):
        raise AssertionError("GIF Compressor should download a GIF file")


@register_case("TC24", "Compression", "tc24_gif_compressor_empty", ROUTES["gif_compressor"])
def tc24_gif_compressor_empty(page, fixtures, timeout_ms):
    locator = button(page, "Compress", exact=True)
    if locator.count() > 0 and locator.is_enabled():
        raise AssertionError('Expected button "Compress" to be disabled')


@register_case("TC25", "Image format conversion", "tc25_convert_image_valid", ROUTES["convert_image"])
def tc25_convert_image_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download", timeout_ms)
    filename = expect_download_on_click(page, "Download", timeout_ms)
    if not any(filename.lower().endswith(ext) for ext in (".jpg", ".jpeg", ".png", ".webp")):
        raise AssertionError("Convert Image should download an image file")


@register_case("TC26", "Image format conversion", "tc26_convert_jpg_valid", ROUTES["convert_jpg"])
def tc26_convert_jpg_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download", timeout_ms)
    filename = expect_download_on_click(page, "Download", timeout_ms)
    if not filename.lower().endswith((".jpg", ".jpeg")):
        raise AssertionError("Convert to JPG should download a JPG file")


@register_case("TC27", "Image format conversion", "tc27_convert_png_valid", ROUTES["convert_png"])
def tc27_convert_png_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_jpg"]])
    assert_text_contains(page, "Convert Image")
    assert_text_contains(page, "Select an image to convert.")
    assert_text_contains(page, "Clear")


@register_case("TC28", "Image format conversion", "tc28_convert_webp_valid", ROUTES["convert_webp"])
def tc28_convert_webp_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download", timeout_ms)
    filename = expect_download_on_click(page, "Download", timeout_ms)
    if not filename.lower().endswith(".webp"):
        raise AssertionError("Convert to WebP should download a WebP file")


@register_case("TC29", "Image format conversion", "tc29_convert_image_invalid", ROUTES["convert_jpg"])
def tc29_convert_image_invalid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["invalid_txt"]])
    assert_button_disabled(page, "Download")
    assert_text_contains(page, "Select an image")


@register_case("TC30", "Image rotation", "tc30_rotate_image_valid", ROUTES["rotate_image"])
def tc30_rotate_image_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download Rotated", timeout_ms)
    button(page, "+90°").click()
    page.get_by_role("checkbox", name="Flip H").check()
    page.get_by_role("checkbox", name="Flip V").check()
    filename = expect_download_on_click(page, "Download Rotated", timeout_ms)
    if not filename.lower().endswith(".png"):
        raise AssertionError("Rotate Image should download a PNG file")


@register_case("TC31", "Image rotation", "tc31_rotate_image_empty", ROUTES["rotate_image"])
def tc31_rotate_image_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Download Rotated")


@register_case("TC32", "Image flipping", "tc32_flip_image_valid", ROUTES["flip_image"])
def tc32_flip_image_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download PNG", timeout_ms)
    page.get_by_role("checkbox", name="Flip Horizontal").check()
    page.get_by_role("checkbox", name="Flip Vertical").check()
    filename = expect_download_on_click(page, "Download PNG", timeout_ms)
    if not filename.lower().endswith(".png"):
        raise AssertionError("Flip Image should download a PNG file")


@register_case("TC33", "Image flipping", "tc33_flip_image_empty", ROUTES["flip_image"])
def tc33_flip_image_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Download PNG")


@register_case("TC34", "Meme generation", "tc34_meme_generator_valid", ROUTES["meme_generator"])
def tc34_meme_generator_valid(page, fixtures, timeout_ms):
    upload_files(page, [fixtures["sample_png"]])
    assert_button_enabled(page, "Download Meme", timeout_ms)
    filename = expect_download_on_click(page, "Download Meme", timeout_ms)
    if not filename.lower().endswith((".png", ".jpg", ".jpeg")):
        raise AssertionError("Meme Generator should download an image file")


@register_case("TC35", "Meme generation", "tc35_meme_generator_empty", ROUTES["meme_generator"])
def tc35_meme_generator_empty(page, fixtures, timeout_ms):
    assert_button_disabled(page, "Download Meme")


@register_case("TC36", "Color picker", "tc36_color_picker_open", ROUTES["color_picker"])
def tc36_color_picker_open(page, fixtures, timeout_ms):
    assert_text_contains(page, "Color Picker")
    assert_text_contains(page, "RGB")
    assert_text_contains(page, "HSV")
    assert_text_contains(page, "HSL")
    assert_text_contains(page, "CMYK")


def run_case(context, base_url: str, case: TestCase, fixtures: dict[str, Path], timeout_ms: int, results_dir: Path) -> dict[str, str]:
    page = context.new_page()
    page.set_default_timeout(timeout_ms)
    url = normalize_path(case.route, base_url)
    screenshot_name = f"{case.name}_fail.png"
    status = "FAIL"

    try:
        page.goto(url, wait_until="domcontentloaded")
        try:
            page.wait_for_load_state("networkidle", timeout=min(timeout_ms, 10000))
        except Exception:
            pass

        case.runner(page, fixtures, timeout_ms)
        status = "PASS"
        screenshot_name = f"{case.name}_pass.png"
        print(f"[{case.tc_id}] {case.feature}: PASS")
    except Exception as exc:
        print(f"[{case.tc_id}] {case.feature}: FAIL - {exc}")
    finally:
        try:
            page.screenshot(path=str(results_dir / screenshot_name), full_page=True)
        except Exception:
            pass
        page.close()

    return {
        "TC_ID": case.tc_id,
        "Feature": case.feature,
        "Status": status,
        "Screenshot": str(results_dir / screenshot_name),
    }


def filter_cases(selected_ids: set[str]) -> list[TestCase]:
    if not selected_ids:
        return TEST_CASES
    return [case for case in TEST_CASES if case.tc_id in selected_ids]


def main() -> int:
    configure_stdout()
    args = parse_args()

    workspace_root = Path(__file__).resolve().parent
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    fixtures = ensure_fixture_files(workspace_root)
    write_csv_header(CSV_PATH)

    selected_ids = {item.strip().upper() for item in args.only.split(",") if item.strip()}
    cases = filter_cases(selected_ids)

    if not cases:
        print("No matching test cases were selected.")
        return 1

    total_pass = 0
    total_fail = 0

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=args.headless, slow_mo=args.slow_mo_ms)
        context = browser.new_context(accept_downloads=True)

        try:
            for case in cases:
                result = run_case(context, args.base_url, case, fixtures, args.timeout_ms, RESULTS_DIR)
                append_csv_row(CSV_PATH, result)
                if result["Status"] == "PASS":
                    total_pass += 1
                else:
                    total_fail += 1
        finally:
            context.close()
            browser.close()

    print("\n========== SUMMARY ==========")
    print(f"Total tests: {len(cases)}")
    print(f"Passed     : {total_pass}")
    print(f"Failed     : {total_fail}")
    print(f"CSV        : {CSV_PATH.resolve()}")
    print(f"Screenshots: {RESULTS_DIR.resolve()}")

    return 0 if total_fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())