"""
convert PPT file to images
"""

import argparse
import os
import subprocess
import tempfile
from concurrent.futures import ThreadPoolExecutor

from pdf2image import convert_from_bytes
from PIL import Image

MAX_WIDTH = 1600  # pick either a max width…
MAX_HEIGHT = 900  # …or a max height


def count_pdf_pages(pdf_path: str) -> int:
    """Use pdfinfo (part of poppler-utils) to count pages."""
    result = subprocess.run(
        ["pdfinfo", pdf_path],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=True,
        text=True,
    )
    for line in result.stdout.splitlines():
        if line.startswith("Pages:"):
            return int(line.split(":")[1].strip())
    raise RuntimeError("Unable to determine number of pages.")


def run_pdftoppm_range(
    pdf_path: str, output_dir: str, thinlinemode: str, start: int, end: int, prefix: str
) -> list[str]:
    output_prefix = os.path.join(output_dir, f"{prefix}_")
    subprocess.run(
        [
            "pdftoppm",
            "-r",
            "300",
            "-png",
            "-thinlinemode",
            thinlinemode,
            "-f",
            str(start),
            "-l",
            str(end),
            pdf_path,
            output_prefix,
        ],
        check=True,
    )
    result = sorted(
        os.path.join(output_dir, f)
        for f in os.listdir(output_dir)
        if f.startswith(f"{prefix}_") and f.endswith(".png")
    )
    return result


def pdf_to_images_with_thinlinemode(
    pdf_bytes: bytes, thinlinemode: str = "none", num_threads: int = 4
) -> list[Image.Image]:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
        temp_pdf.write(pdf_bytes)
        temp_pdf_path = temp_pdf.name

    total_pages = count_pdf_pages(temp_pdf_path)
    pages_per_thread = max(1, total_pages // num_threads)

    with tempfile.TemporaryDirectory() as temp_dir:
        futures = []
        with ThreadPoolExecutor(max_workers=num_threads) as executor:
            for i in range(0, total_pages, pages_per_thread):
                start = i + 1
                end = min(start + pages_per_thread - 1, total_pages)
                prefix = f"range_{start}_{end}"
                futures.append(
                    executor.submit(
                        run_pdftoppm_range,
                        temp_pdf_path,
                        temp_dir,
                        thinlinemode,
                        start,
                        end,
                        prefix,
                    )
                )

        # Gather and flatten results
        all_image_paths = []
        for future in futures:
            all_image_paths.extend(future.result())

        images: list[Image.Image] = []
        for img_path in sorted(all_image_paths, key=lambda p: int(p.split("-")[-1].split(".")[0])):
            with open(img_path, "rb") as img_file:
                img = Image.open(img_file)
                img.load()
                images.append(img)

    os.remove(temp_pdf_path)
    return images


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert PPTX file to images.")
    parser.add_argument(
        "--input",
        type=str,
        required=True,
        help="Path to the input PPTX file (or directory containing PPTX file).",
    )
    parser.add_argument(
        "--output",
        type=str,
        nargs="?",
        default="ppt-preview",
        help="Directory to save output images (default: ppt-preview)",
    )
    args = parser.parse_args()

    img_format = "png"
    pptfile_name = args.input
    out_dir = args.output

    filename_base = os.path.basename(pptfile_name)
    filename_bare = os.path.splitext(filename_base)[0]

    # soffice --headless --convert-to pdf demo.pptx
    # convert pptx to PDF
    command_list = ["soffice", "--headless", "--convert-to", "pdf", pptfile_name]
    subprocess.run(command_list)

    pdffile_name = filename_bare + ".pdf"
    with open(pdffile_name, "rb") as f:
        pdf_bytes = f.read()

    images = convert_from_bytes(pdf_bytes, dpi=300, thread_count=8)

    if not os.path.exists(out_dir):
        os.mkdir(out_dir)

    def resize_and_save(args: tuple[int, Image.Image]) -> None:
        i, img = args
        w, h = img.size
        scale = min(MAX_WIDTH / w, MAX_HEIGHT / h)
        im_name = os.path.join(out_dir, f"slide-{i}.{img_format}")
        new_size = (int(w * scale), int(h * scale))
        resized_img = img.resize(new_size, Image.Resampling.LANCZOS)
        resized_img.save(im_name)

    with ThreadPoolExecutor(max_workers=8) as executor:
        list(executor.map(resize_and_save, enumerate(images)))


if __name__ == "__main__":
    main()
