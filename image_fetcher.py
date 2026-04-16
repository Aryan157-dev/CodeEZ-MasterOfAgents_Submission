import os
import re
import requests
from io import BytesIO

PEXELS_API_KEY = os.environ.get("PEXELS_API_KEY", "")
CACHE = {}  # simple in-memory cache to avoid duplicate requests

# ── Keyword extractor ─────────────────────────────────────────────────────
def extract_keyword(title, doc_title=""):
    """Extract best search keyword combining slide title + doc context."""
    generic = {"overview", "introduction", "summary", "conclusion",
               "key", "analysis", "slide", "section", "findings",
               "background", "methodology", "scope", "objectives",
               "comparison", "trends", "strategies", "recommendations"}

    def meaningful_words(text):
        words = re.sub(r'[^a-zA-Z0-9\s]', '', text.lower()).split()
        return [w for w in words if w not in generic and len(w) > 3]

    slide_words = meaningful_words(title)
    doc_words   = meaningful_words(doc_title)

    # Combine: 1-2 words from slide title + 1 word from doc title for context
    combined = []
    if slide_words:
        combined.extend(slide_words[:2])
    if doc_words:
        # Add doc context word if it's not already in combined
        for w in doc_words[:2]:
            if w not in combined:
                combined.append(w)
                break

    if combined:
        return " ".join(combined[:3])

    return doc_title[:30] if doc_title else "business professional"


def fetch_image(keyword, orientation="landscape", size="large"):
    """
    Fetch a relevant image from Pexels.
    Returns image bytes or None if failed.
    """
    if not PEXELS_API_KEY:
        return None

    # Check cache
    cache_key = keyword.lower().strip()
    if cache_key in CACHE:
        return BytesIO(CACHE[cache_key])

    try:
        headers = {"Authorization": PEXELS_API_KEY}
        params = {
            "query": keyword,
            "per_page": 3,
            "orientation": orientation,
            "size": size
        }
        r = requests.get("https://api.pexels.com/v1/search",
                         headers=headers, params=params, timeout=8)

        if r.status_code != 200:
            print(f"   ⚠️  Pexels API error {r.status_code} for '{keyword}'")
            return None

        data = r.json()
        photos = data.get("photos", [])
        if not photos:
            print(f"   ⚠️  No images found for '{keyword}'")
            return None

        # Pick first photo
        img_url = photos[0]["src"]["large"]
        img_r = requests.get(img_url, timeout=10)
        if img_r.status_code == 200:
            img_bytes = img_r.content  # store raw bytes, not BytesIO
            CACHE[cache_key] = img_bytes
            print(f"   🖼️  Image fetched for '{keyword}'")
            return BytesIO(img_bytes)

    except Exception as e:
        print(f"   ⚠️  Image fetch failed for '{keyword}': {e}")

    return None


def fetch_title_image(doc_title):
    """Fetch a high-quality image specifically for the title slide."""
    keyword = extract_keyword(doc_title, doc_title)
    return fetch_image(keyword, orientation="landscape", size="large2x")