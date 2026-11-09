#!/usr/bin/env python3
"""Prepare web and source-document images for the Mission investment deck."""

from __future__ import annotations

import json
import subprocess
import tempfile
from pathlib import Path

import requests
from PIL import Image, ImageEnhance


ROOT = Path(__file__).resolve().parents[1]
ASSET_DIR = ROOT / "assets" / "generated"
UPLOAD_DIR = Path(
    "/home/ubuntu/.cursor/projects/agent-repos-wenjian/uploads"
)
EV_PDF = UPLOAD_DIR / "_____Mission_EV_Assembly_Proposal_CN_2026_04cd.pdf"
BUSINESS_PDF = (
    UPLOAD_DIR
    / "12458_Subdivision_Mission__Business_Plan_CN_2026_Final_1fc0.pdf"
)

COMMONS_IMAGES = {
    "mission_fraser_river.jpg": "File:FraserRiver.jpg",
    "vancouver_centerm.jpg": (
        "File:Vancouver (BC, Canada), Vancouver Centerm Terminal -- "
        "2022 -- 1848.jpg"
    ),
    "distribution_centre.jpg": (
        "File:Distribution Center Rhenus Logistics Ekkersrijt.jpg"
    ),
}


def download_commons(title: str, destination: Path, width: int = 2400) -> dict:
    response = requests.get(
        "https://commons.wikimedia.org/w/api.php",
        params={
            "action": "query",
            "format": "json",
            "prop": "imageinfo",
            "iiprop": "url|extmetadata",
            "iiurlwidth": width,
            "titles": title,
        },
        headers={"User-Agent": "MissionInvestmentDeck/1.0"},
        timeout=60,
    )
    response.raise_for_status()
    page = next(iter(response.json()["query"]["pages"].values()))
    info = page["imageinfo"][0]
    image_url = info.get("thumburl", info["url"])
    image_response = requests.get(
        image_url,
        headers={"User-Agent": "MissionInvestmentDeck/1.0"},
        timeout=60,
    )
    image_response.raise_for_status()
    destination.write_bytes(image_response.content)
    return {
        "title": title,
        "description_url": info["descriptionurl"],
        "image_url": image_url,
    }


def render_page(pdf: Path, page: int, destination: Path, dpi: int = 180) -> None:
    with tempfile.TemporaryDirectory() as temp_dir:
        prefix = Path(temp_dir) / "page"
        subprocess.run(
            [
                "pdftoppm",
                "-f",
                str(page),
                "-singlefile",
                "-jpeg",
                "-r",
                str(dpi),
                str(pdf),
                str(prefix),
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        Image.open(prefix.with_suffix(".jpg")).convert("RGB").save(
            destination, quality=94
        )


def crop_relative(
    source: Path,
    destination: Path,
    box: tuple[float, float, float, float],
    *,
    contrast: float = 1.0,
) -> None:
    image = Image.open(source).convert("RGB")
    width, height = image.size
    crop_box = (
        round(box[0] * width),
        round(box[1] * height),
        round(box[2] * width),
        round(box[3] * height),
    )
    cropped = image.crop(crop_box)
    if contrast != 1.0:
        cropped = ImageEnhance.Contrast(cropped).enhance(contrast)
    cropped.save(destination, quality=94)


def main() -> None:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    sources: dict[str, object] = {
        "generated_from_user_documents": [],
        "web_images": [],
    }

    for filename, title in COMMONS_IMAGES.items():
        metadata = download_commons(title, ASSET_DIR / filename)
        sources["web_images"].append({"file": filename, **metadata})

    page_specs = {
        "facility_aerial.jpg": (EV_PDF, 7, (0.043, 0.198, 0.958, 0.718)),
        "facility_entry.jpg": (EV_PDF, 42, (0.441, 0.0, 1.0, 1.0)),
        "facility_loading.jpg": (EV_PDF, 22, (0.043, 0.198, 0.958, 0.718)),
        "visitor_experience.jpg": (EV_PDF, 10, (0.043, 0.198, 0.958, 0.765)),
        "factory_gallery.jpg": (EV_PDF, 12, (0.043, 0.198, 0.958, 0.765)),
        "light_manufacturing.jpg": (EV_PDF, 20, (0.043, 0.198, 0.958, 0.718)),
        "mixed_use.jpg": (EV_PDF, 29, (0.043, 0.198, 0.958, 0.765)),
    }
    for filename, (pdf, page, box) in page_specs.items():
        rendered = ASSET_DIR / f"_page_{page}.jpg"
        if not rendered.exists():
            render_page(pdf, page, rendered)
        crop_relative(rendered, ASSET_DIR / filename, box, contrast=1.03)
        sources["generated_from_user_documents"].append(
            {"file": filename, "document": pdf.name, "page": page}
        )

    business_page = ASSET_DIR / "_business_page_31.jpg"
    render_page(BUSINESS_PDF, 31, business_page)
    crop_relative(
        business_page,
        ASSET_DIR / "parcel_zoning_map.jpg",
        (0.012, 0.012, 0.828, 0.985),
        contrast=1.08,
    )
    sources["generated_from_user_documents"].append(
        {
            "file": "parcel_zoning_map.jpg",
            "document": BUSINESS_PDF.name,
            "page": 31,
        }
    )

    (ASSET_DIR / "sources.json").write_text(
        json.dumps(sources, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
