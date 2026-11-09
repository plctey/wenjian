"""Download the photos listed in photo_manifest.json from Wikimedia Commons and record attribution."""
import json
import time
import urllib.parse
import urllib.request
from pathlib import Path

from PIL import Image

from search_images import API, UA, strip

HERE = Path(__file__).parent
PHOTOS = HERE / "assets" / "photos"
PHOTOS.mkdir(parents=True, exist_ok=True)
MAX_WIDTH = 2000


def info(title):
    q = urllib.parse.urlencode({
        "action": "query", "titles": title, "prop": "imageinfo", "format": "json",
        "iiprop": "url|size|extmetadata", "iiurlwidth": MAX_WIDTH,
    })
    with urllib.request.urlopen(urllib.request.Request(f"{API}?{q}", headers=UA), timeout=30) as r:
        page = next(iter(json.load(r)["query"]["pages"].values()))
    return page["imageinfo"][0]


def main():
    manifest = json.loads((HERE / "photo_manifest.json").read_text())
    credits = []
    for item in manifest:
        ii = info(item["title"])
        md = ii.get("extmetadata", {})
        url = ii.get("thumburl") if ii["width"] > MAX_WIDTH else ii["url"]
        dest = PHOTOS / f"{item['key']}.jpg"
        if not dest.exists():
            for attempt in range(4):
                try:
                    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60) as r:
                        dest.write_bytes(r.read())
                    break
                except Exception as exc:
                    print("retry", item["key"], exc)
                    time.sleep(4 * 2 ** attempt)
            im = Image.open(dest).convert("RGB")
            im.thumbnail((1920, 1920))
            im.save(dest, quality=85, optimize=True)
        credits.append({
            "key": item["key"],
            "title": item["title"].removeprefix("File:"),
            "artist": strip(md.get("Artist", {}).get("value", "")).replace("&amp;", "&"),
            "license": md.get("LicenseShortName", {}).get("value", ""),
            "page": ii.get("descriptionurl"),
        })
        print("ok", item["key"])
    (HERE / "photo_credits.json").write_text(json.dumps(credits, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
