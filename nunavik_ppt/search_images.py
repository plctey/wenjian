"""Search Wikimedia Commons and list candidate images with license info.

Usage: python3 search_images.py "query" [limit]
"""
import json
import re
import sys
import urllib.parse
import urllib.request

API = "https://commons.wikimedia.org/w/api.php"
UA = {"User-Agent": "nunavik-ppt-builder/1.0 (presentation image research)"}


def api(params):
    url = API + "?" + urllib.parse.urlencode({**params, "format": "json"})
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
        return json.load(r)


def strip(html):
    return re.sub(r"<[^>]+>", "", html or "").strip()


def search(query, limit=12):
    data = api({
        "action": "query", "generator": "search", "gsrsearch": f"filetype:bitmap {query}",
        "gsrnamespace": 6, "gsrlimit": limit, "prop": "imageinfo",
        "iiprop": "url|size|extmetadata", "iiurlwidth": 400,
    })
    out = []
    for p in sorted(data.get("query", {}).get("pages", {}).values(), key=lambda p: p.get("index", 0)):
        ii = p["imageinfo"][0]
        md = ii.get("extmetadata", {})
        out.append({
            "title": p["title"],
            "w": ii["width"], "h": ii["height"],
            "license": md.get("LicenseShortName", {}).get("value", ""),
            "artist": strip(md.get("Artist", {}).get("value", "")),
            "thumb": ii.get("thumburl"),
            "url": ii["url"],
            "page": ii.get("descriptionurl"),
        })
    return out


if __name__ == "__main__":
    q = sys.argv[1]
    n = int(sys.argv[2]) if len(sys.argv) > 2 else 12
    for r in search(q, n):
        print(json.dumps(r, ensure_ascii=False))
