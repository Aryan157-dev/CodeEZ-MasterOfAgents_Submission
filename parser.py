import re


def clean_text(text):
    """Remove markdown links, extra whitespace, reference numbers"""
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'\[\d+\]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def parse_markdown(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    result = {
        "title": "",
        "subtitle": "",
        "executive_summary": "",
        "sections": [],
        "tables": [],
        "references": [],
        "key_stats": [],          # NEW — for kpi_stats layout
        "timeline_events": [],    # NEW — for timeline layout
        "has_comparisons": False, # NEW — for two_col_compare layout
    }

    lines = content.split('\n')
    current_section = None
    current_content = []

    for line in lines:
        # Title (H1)
        if line.startswith('# ') and not result["title"]:
            result["title"] = line[2:].strip()

        # Subtitle (### under title, before any sections)
        elif line.startswith('### ') and not result["subtitle"] and not result["sections"]:
            result["subtitle"] = line[4:].strip()

        # H2 sections
        elif line.startswith('## '):
            if current_section:
                current_section["content"] = '\n'.join(current_content).strip()
                result["sections"].append(current_section)
            section_title = line[3:].strip()
            if "executive summary" in section_title.lower():
                current_section = {"title": section_title, "content": "", "level": 2, "is_executive_summary": True}
            elif "reference" in section_title.lower():
                current_section = {"title": section_title, "content": "", "level": 2, "is_references": True}
            else:
                current_section = {"title": section_title, "content": "", "level": 2}
            current_content = []

        # H3 subsections
        elif line.startswith('### '):
            current_content.append(line)

        # Tables
        elif line.startswith('|'):
            current_content.append(line)

        else:
            current_content.append(line)

    # Last section
    if current_section:
        current_section["content"] = '\n'.join(current_content).strip()
        result["sections"].append(current_section)

    # Executive summary text
    for section in result["sections"]:
        if section.get("is_executive_summary"):
            result["executive_summary"] = section["content"]
            break

    # Tables
    result["tables"] = extract_tables(content)

    # NEW extractions
    result["key_stats"]       = extract_key_stats(content)
    result["timeline_events"] = extract_timeline_events(content)
    result["has_comparisons"] = detect_comparisons(result["sections"])

    return result


def extract_tables(content):
    """Extract all markdown tables with their titles."""
    tables = []
    lines = content.split('\n')
    i = 0
    while i < len(lines):
        if lines[i].startswith('|') and i + 1 < len(lines) and '---' in lines[i + 1]:
            # Title = nearest non-empty, non-table line above
            title = ""
            if i > 0:
                j = i - 1
                while j >= 0 and lines[j].strip() == '':
                    j -= 1
                if j >= 0 and not lines[j].startswith('|'):
                    title = lines[j].strip()
                    if title.lower().startswith('title:'):
                        title = title[6:].strip()
                    # Strip markdown heading markers
                    title = re.sub(r'^#+\s*', '', title).strip()

            headers = [h.strip() for h in lines[i].split('|') if h.strip()]
            rows = []
            i += 2  # skip header + separator
            while i < len(lines) and lines[i].startswith('|'):
                row = [clean_text(c.strip()) for c in lines[i].split('|') if c.strip()]
                if row:
                    rows.append(row)
                i += 1
            tables.append({"title": title, "headers": headers, "rows": rows})
        else:
            i += 1
    return tables


def extract_key_stats(content):
    """
    Extract key numerical stats for kpi_stats slides.
    Returns list of dicts: {label, value, context}
    """
    stats = []
    seen  = set()

    patterns = [
        # $X billion / $X million / $XB / $XM
        (r'\$[\d,.]+\s*(?:billion|million|B|M)\b', "Financial"),
        # X% or X.X%
        (r'\b\d{1,3}(?:\.\d+)?%', "Percentage"),
        # plain large numbers with context words
        (r'\b\d{2,4}\s+(?:acquisitions?|companies|employees|professionals?|firms?)\b', "Count"),
        # revenue / investment figures
        (r'\b(?:revenue|investment|bookings?|spending)\s+of\s+\$[\d,.]+\s*(?:billion|million|B|M)?\b', "Financial"),
    ]

    lines = content.split('\n')
    for line in lines:
        clean = clean_text(line)
        for pattern, category in patterns:
            matches = re.findall(pattern, clean, re.IGNORECASE)
            for match in matches:
                match = match.strip()
                if match in seen:
                    continue
                seen.add(match)
                # Try to grab context: a few words around the match
                ctx_match = re.search(
                    r'(.{0,40})' + re.escape(match) + r'(.{0,60})',
                    clean, re.IGNORECASE
                )
                context = ""
                if ctx_match:
                    context = (ctx_match.group(1) + ctx_match.group(2)).strip()
                    context = re.sub(r'\s+', ' ', context)[:80]

                stats.append({
                    "value": match,
                    "label": category,
                    "context": context
                })

    # Deduplicate and return top 6
    return stats[:6]


def extract_timeline_events(content):
    """
    Extract year-based timeline events.
    Returns list of dicts: {year, event}
    """
    events = []
    seen_years = set()

    # Match lines containing a 4-digit year (1990-2030) followed by content
    year_pattern = re.compile(r'\b(20\d{2}|19\d{2})\b')

    lines = content.split('\n')
    for line in lines:
        clean = clean_text(line)
        if not clean or len(clean) < 10:
            continue
        years = year_pattern.findall(clean)
        for year in years:
            if year in seen_years:
                continue
            # Skip lines that are just table rows with many years
            if clean.count(year) > 2:
                continue
            seen_years.add(year)
            # Remove the year from the line to get the event description
            event = re.sub(r'\b' + year + r'\b', '', clean).strip()
            event = re.sub(r'^[\s:–\-|]+', '', event).strip()
            event = re.sub(r'\s+', ' ', event)
            if len(event) > 10:
                events.append({"year": year, "event": event[:120]})

    # Sort chronologically
    events.sort(key=lambda x: x["year"])
    return events[:8]


def detect_comparisons(sections):
    """
    Returns True if any section title suggests a comparison/contrast layout.
    """
    comparison_keywords = [
        "vs", "versus", "comparison", "benchmark", "peer",
        "success", "challenge", "pros", "cons", "strength",
        "weakness", "opportunity", "risk", "advantage", "disadvantage",
        "before", "after", "difference", "contrast"
    ]
    for section in sections:
        title_lower = section["title"].lower()
        if any(kw in title_lower for kw in comparison_keywords):
            return True
        # Also check content for comparison patterns
        content_lower = section.get("content", "").lower()
        compare_count = sum(1 for kw in ["success", "challenge", "advantage", "disadvantage", "strength", "weakness"] if kw in content_lower)
        if compare_count >= 2:
            return True
    return False


# ── Test ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    result = parse_markdown("Accenture Tech Acquisition Analysis.md")

    print("Title:", result["title"])
    print("Subtitle:", result["subtitle"][:80])
    print("Sections:", len(result["sections"]))
    print("Tables:", len(result["tables"]))

    print("\nKey Stats found:")
    for s in result["key_stats"]:
        print(f"  [{s['label']}] {s['value']} — {s['context'][:60]}")

    print("\nTimeline Events found:")
    for e in result["timeline_events"][:5]:
        print(f"  {e['year']}: {e['event'][:70]}")

    print("\nHas comparisons:", result["has_comparisons"])

    print("\nSections:")
    for s in result["sections"]:
        print(f"  - {s['title'][:60]}")

    print("\nTables:")
    for t in result["tables"]:
        print(f"  - {t['title'][:60]} ({len(t['rows'])} rows)")