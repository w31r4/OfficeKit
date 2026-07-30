#!/usr/bin/env python3
# Copyright (c) w31r4. All rights reserved.
import argparse
from glob import glob
import json
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from os import makedirs, replace
from os.path import abspath, basename, dirname, exists, expanduser, join, splitext
from typing import Sequence, cast
from zipfile import ZipFile

SCRIPT_DIR = dirname(__file__)
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from runtime_tools import node_binary, runtime_binary, runtime_env  # noqa: E402

EMU_PER_INCH: int = 914_400
PRESENTATION_EXTS = (".pptx", ".ppsx", ".potx", ".pptm", ".ppsm", ".potm")


def _run_poppler(command: list[str], label: str) -> subprocess.CompletedProcess[str]:
    """Run one required Poppler command with an actionable failure.

    These helpers intentionally call the bundled/runtime-resolved Poppler
    binaries directly.  A third-party Python wrapper is not part of this
    skill's runtime contract.
    """

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            check=False,
            text=True,
            env=runtime_env(),
        )
    except FileNotFoundError as error:
        raise RuntimeError(
            f"{label} is unavailable: expected `{command[0]}` on PATH. "
            "Expose the required Poppler command; this helper does not install providers."
        ) from error
    if result.returncode != 0:
        details = (result.stderr or result.stdout or "no output").strip()
        raise RuntimeError(f"{label} failed ({result.returncode}): {details}")
    return result


def calc_dpi_via_ooxml(input_path: str, max_w_px: int, max_h_px: int) -> int:
    """Calculate DPI from OOXML `ppt/presentation.xml` slide size (cx/cy in EMUs)."""
    with ZipFile(input_path, "r") as zf:
        xml = zf.read("ppt/presentation.xml")
    root = ET.fromstring(xml)
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    sld_sz = root.find("p:sldSz", ns)
    if sld_sz is None:
        raise RuntimeError("Slide size not found in presentation.xml")
    cx = int(sld_sz.get("cx") or 0)
    cy = int(sld_sz.get("cy") or 0)
    if cx <= 0 or cy <= 0:
        raise RuntimeError("Invalid slide size values in presentation.xml")
    width_in = cx / EMU_PER_INCH
    height_in = cy / EMU_PER_INCH
    return round(min(max_w_px / width_in, max_h_px / height_in))


def calc_dpi_via_pdf(input_path: str, max_w_px: int, max_h_px: int) -> int:
    """Compute DPI from PDF page size.

    For OOXML presentation inputs, use the presentation slide size directly.
    For PDFs, use the PDF directly.
    """
    is_pdf = input_path.lower().endswith(".pdf")
    if not is_pdf:
        if input_path.lower().endswith(PRESENTATION_EXTS):
            return calc_dpi_via_ooxml(input_path, max_w_px, max_h_px)
        raise RuntimeError("DPI computation is supported for PDF and OOXML presentation inputs.")

    info = _run_poppler(
        [runtime_binary("pdfinfo"), input_path],
        "Poppler pdfinfo",
    ).stdout
    size_val = None
    for line in info.splitlines():
        label, separator, value = line.partition(":")
        if separator and label.strip().lower() == "page size":
            size_val = value.strip()
            break
    if not size_val:
        for line in info.splitlines():
            if "size" in line.lower() and "pts" in line.lower():
                _, _, value = line.partition(":")
                size_val = value.strip() if value else line.strip()
                break
    if not isinstance(size_val, str):
        raise RuntimeError("Failed to read PDF page size for DPI computation.")

    def _parse_page_size_to_pts(s: str) -> tuple[float, float]:
        # Common formats from poppler/pdfinfo:
        # - "612 x 792 pts (letter)"
        # - "595.276 x 841.89 pts (A4)"
        # - sometimes inches: "8.5 x 11 in"
        m_pts = re.search(
            r"([0-9]+(?:\.[0-9]+)?)\s*x\s*([0-9]+(?:\.[0-9]+)?)\s*pts\b",
            s,
        )
        if m_pts:
            return float(m_pts.group(1)), float(m_pts.group(2))
        m_in = re.search(
            r"([0-9]+(?:\.[0-9]+)?)\s*x\s*([0-9]+(?:\.[0-9]+)?)\s*in\b",
            s,
        )
        if m_in:
            w_in = float(m_in.group(1))
            h_in = float(m_in.group(2))
            return w_in * 72.0, h_in * 72.0
        # Sometimes poppler returns without an explicit unit; treat as points.
        m = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*x\s*([0-9]+(?:\.[0-9]+)?)\b", s)
        if m:
            return float(m.group(1)), float(m.group(2))
        raise RuntimeError(f"Unrecognized PDF page size format: {s!r}")

    width_pts, height_pts = _parse_page_size_to_pts(size_val)
    width_in = width_pts / 72.0
    height_in = height_pts / 72.0
    if width_in <= 0 or height_in <= 0:
        raise RuntimeError("Invalid PDF page size values.")
    return round(min(max_w_px / width_in, max_h_px / height_in))


def run_cmd_no_check(cmd: list[str]) -> None:
    subprocess.run(
        cmd,
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        env=runtime_env(),
    )


def _render_presentation_with_artifact_tool(
    input_path: str, out_dir: str, dpi: int
) -> Sequence[str]:
    scale = max(dpi / 96.0, 0.01)
    with tempfile.TemporaryDirectory(prefix="artifact_tool_workspace_") as workspace:
        proc = subprocess.run(
            [
                node_binary(),
                join(SCRIPT_DIR, "render_presentation.mjs"),
                "--input",
                input_path,
                "--output_dir",
                out_dir,
                "--scale",
                f"{scale:.6f}",
                "--workspace",
                workspace,
            ],
            capture_output=True,
            text=True,
            check=False,
            env=runtime_env(),
        )
    if proc.returncode != 0:
        details = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(
            "Failed to render presentation with artifact-tool."
            + (f"\n{details}" if details else "")
        )
    payload = json.loads(proc.stdout)
    return cast(Sequence[str], payload["paths"])


def convert_to_pdf(
    pptx_path: str,
    user_profile: str,
    convert_tmp_dir: str,
    stem: str,
) -> str:
    pdf_path = join(convert_tmp_dir, f"{stem}.pdf")
    if pptx_path.lower().endswith(".pdf"):
        return pptx_path
    if not pptx_path.lower().endswith(PRESENTATION_EXTS):
        return ""

    image_dir = join(convert_tmp_dir, stem + "_pdf_pages")
    paths = list(_render_presentation_with_artifact_tool(pptx_path, image_dir, dpi=96))
    if not paths:
        return ""

    # This legacy assembly helper is the only path in this module that needs
    # Pillow.  PPTX/PDF PNG rendering itself is standard-library + Poppler.
    from PIL import Image

    images = []
    try:
        for path in paths:
            images.append(Image.open(path).convert("RGB"))
        first, *rest = images
        first.save(pdf_path, "PDF", save_all=True, append_images=rest)
    finally:
        for image in images:
            image.close()
    return pdf_path if exists(pdf_path) else ""


def rasterize(
    input_path: str,
    out_dir: str,
    dpi: int,
) -> Sequence[str]:
    """Rasterise PPTX/PDF to PNG files placed in out_dir and return the image paths."""
    makedirs(out_dir, exist_ok=True)
    input_path = abspath(input_path)

    if input_path.lower().endswith(PRESENTATION_EXTS):
        return _render_presentation_with_artifact_tool(input_path, out_dir, dpi)

    if not input_path.lower().endswith(".pdf"):
        raise RuntimeError("Rasterization is supported for PDF and OOXML presentation inputs.")

    # Stage beneath the destination so promotion remains an atomic rename even
    # when the caller chose a mounted output directory.
    with tempfile.TemporaryDirectory(prefix=".officekit-poppler-", dir=out_dir) as stage_dir:
        prefix = join(stage_dir, "slide")
        _run_poppler(
            [runtime_binary("pdftoppm"), "-png", "-r", str(dpi), input_path, prefix],
            "Poppler pdftoppm",
        )
        slides = []
        for src_path in glob(join(stage_dir, "slide-*.png")):
            match = re.fullmatch(r"slide-(\d+)\.png", basename(src_path))
            if not match:
                continue
            slide_num = int(match.group(1))
            dst_path = join(out_dir, f"slide-{slide_num}.png")
            replace(src_path, dst_path)
            slides.append((slide_num, dst_path))
        if not slides:
            raise RuntimeError("Poppler pdftoppm completed without producing slide PNGs.")
    slides.sort(key=lambda t: t[0])
    final_paths = [path for _, path in slides]
    return final_paths


def main() -> None:
    parser = argparse.ArgumentParser(description="Render slides to images.")
    parser.add_argument(
        "input_path",
        type=str,
        help="Path to the input PowerPoint or PDF file.",
    )
    parser.add_argument(
        "--output_dir",
        type=str,
        default=None,
        help=(
            "Output directory for the rendered images. "
            "Defaults to a folder next to the input named after the input file (without extension)."
        ),
    )
    parser.add_argument(
        "--width",
        type=int,
        default=1600,
        help=(
            "Approximate maximum width in pixels after isotropic scaling (default 1600). "
            "The actual value may exceed slightly."
        ),
    )
    parser.add_argument(
        "--height",
        type=int,
        default=900,
        help=(
            "Approximate maximum height in pixels after isotropic scaling (default 900). "
            "The actual value may exceed slightly."
        ),
    )
    args = parser.parse_args()

    input_path = abspath(expanduser(args.input_path))
    out_dir = abspath(expanduser(args.output_dir)) if args.output_dir else splitext(input_path)[0]
    if input_path.lower().endswith((".pptx", ".ppsx", ".potx", ".pptm", ".ppsm", ".potm")):
        dpi = calc_dpi_via_ooxml(input_path, args.width, args.height)
    else:
        dpi = calc_dpi_via_pdf(input_path, args.width, args.height)
    rasterize(input_path, out_dir, dpi)
    print("Slides rendered to " + out_dir)


if __name__ == "__main__":
    main()
