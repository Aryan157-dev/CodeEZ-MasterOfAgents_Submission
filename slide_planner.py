from langchain_groq import ChatGroq
from langchain_core.messages import HumanMessage
import json
import re

# Valid layout types the builder supports
VALID_LAYOUT_TYPES = {
    "title", "executive_summary", "split_panel", "timeline",
    "kpi_stats", "kpi_visual", "two_col_compare", "grid_4col", "data_table",
    "chart", "conclusion"
}

# Fallback mapping: if LLM gives old slide_type, map to layout_type
LEGACY_TYPE_MAP = {
    "content": "split_panel",
    "executive_summary": "executive_summary",
    "title": "title",
    "data_table": "data_table",
    "chart": "chart",
    "conclusion": "conclusion"
}


def sanitize_slide(slide, index):
    """Ensure every slide has valid fields. Fixes bad LLM output."""
    slide["slide_number"] = slide.get("slide_number", index + 1)
    slide["title"] = str(slide.get("title", "Untitled"))[:80]
    slide["key_points"] = slide.get("key_points", [])
    slide["use_table"] = slide.get("use_table", None)
    slide["speaker_notes"] = slide.get("speaker_notes", "")
    slide["metadata"] = slide.get("metadata", {})

    # Resolve layout_type
    layout = slide.get("layout_type", "").strip().lower()
    if layout not in VALID_LAYOUT_TYPES:
        # Try to map from legacy slide_type
        old_type = slide.get("slide_type", "").strip().lower()
        layout = LEGACY_TYPE_MAP.get(old_type, "split_panel")
    slide["layout_type"] = layout

    # Ensure key_points is a list of strings
    if not isinstance(slide["key_points"], list):
        slide["key_points"] = []
    slide["key_points"] = [str(p) for p in slide["key_points"] if p]

    return slide


def plan_slides(parsed_doc, api_key):
    """Use LLM to plan 10-15 slides with rich layout intelligence."""

    llm = ChatGroq(
        model="llama-3.3-70b-versatile",
        api_key=api_key
    )

    # Build document context for the LLM
    sections_list = "\n".join([
        f"- {s['title']}" for s in parsed_doc["sections"]
        if not s.get("is_references")
    ])

    tables_list = "\n".join([
        f"- {t['title']} ({len(t['rows'])} rows, columns: {', '.join(t['headers'][:4])})"
        for t in parsed_doc["tables"]
    ]) if parsed_doc["tables"] else "None"

    exec_summary = parsed_doc.get("executive_summary", "")[:800]

    # Extract key stats for KPI slide hint
    stats = parsed_doc.get("key_stats", [])
    stats_hint = ", ".join([s["value"] for s in stats[:6]]) if stats else "None found"

    # Check for timeline signals in section titles
    timeline_hint = any(
        any(kw in s["title"].lower() for kw in ["timeline", "history", "evolution", "milestones", "journey", "year", "phase"])
        for s in parsed_doc["sections"]
    )

    # Check for comparison signals
    compare_hint = any(
        any(kw in s["title"].lower() for kw in ["vs", "comparison", "benchmark", "peer", "success", "challenge", "pros", "cons", "versus"])
        for s in parsed_doc["sections"]
    )

    prompt = f"""You are a professional PowerPoint presentation designer. Plan a presentation with 10-13 slides.

DOCUMENT TITLE: {parsed_doc["title"]}
SUBTITLE: {parsed_doc.get("subtitle", "")}

EXECUTIVE SUMMARY:
{exec_summary}

SECTIONS IN DOCUMENT:
{sections_list}

TABLES AVAILABLE:
{tables_list}

KEY STATISTICS FOUND: {stats_hint}
TIMELINE CONTENT DETECTED: {timeline_hint}
COMPARISON CONTENT DETECTED: {compare_hint}

=== LAYOUT TYPE GUIDE (choose the best fit for each slide) ===

"title"
  → Use for: Slide 1 only. The opening title slide.

"executive_summary"  
  → Use for: Slide 2 only. 5 key insight cards with Header: Description format.

"kpi_stats"
  → Use for: Slides with KEY NUMBERS or metrics (e.g. $6.6B, 326 acquisitions, 47%).
  → key_points: each point is a stat in format "LABEL: VALUE: context" e.g. "Total Acquisitions: 326: completed 2020-2025"
  → Use when: stats_hint has real numbers, or section is about financial results/metrics.

"kpi_visual"
  → Use for: A visual dashboard slide with circular progress rings showing key metrics.
  → key_points: same format as kpi_stats "LABEL: VALUE: context"
  → Use ONCE per presentation when strong numerical stats exist.
  → More visual and impactful than kpi_stats.

"timeline"
  → Use for: Historical progression, milestones, phases, year-by-year events.
  → key_points: each point is "YEAR/PHASE: Event description" e.g. "2020: Began aggressive AI acquisition push"
  → Use when: timeline_hint is True, or section covers evolution/history.

"two_col_compare"
  → Use for: Successes vs Challenges, Pros vs Cons, Before vs After, Region A vs Region B.
  → key_points: first half are LEFT column points, second half are RIGHT column points (split evenly).
  → metadata: {{"left_label": "Successes", "right_label": "Challenges"}}
  → Use when: compare_hint is True, or section has opposing concepts.

"grid_4col"
  → Use for: 4 distinct categories, pillars, domains, or strategies displayed as a grid.
  → key_points: exactly 4 items, each is "COLUMN_HEADER: description"
  → Use when: section has 4 clear themes (e.g. Geographic distribution with 4 regions).

"split_panel"
  → Use for: General content sections with 3-5 bullet points. Left panel = bold context, right = points.
  → - key_points: list of EXACTLY 5 bullet points, each 10-15 words, format as "Bold Label: detailed description with specific facts"
  → This is the DEFAULT layout for most content slides.

"data_table"
  → Use for: Slides that should show a table from the document.
  → Set use_table to the table title.

"chart"
  → Use for: Data visualization from a table (bar/line chart).
  → use_table: MUST be copied EXACTLY verbatim from the TABLES AVAILABLE list above. Do not paraphrase or describe it. Copy the exact table title string character by character.

"conclusion"
  → Use for: The final slide. Key takeaways with → arrows.

=== SLIDE STRUCTURE RULES ===
- Slide 1: title
- Slide 2: executive_summary  
- Slides 3-N: Mix of split_panel, kpi_stats, timeline, two_col_compare, grid_4col, data_table, chart
- Last slide: conclusion
- If tables exist: include at least one data_table AND one chart slide
- If key stats exist: include one kpi_stats slide
- If timeline detected: include one timeline slide
- If comparison detected: include one two_col_compare slide
- Total slides: 10-13
- For data_table and chart slides: use_table MUST be an exact copy of one of the table titles from TABLES AVAILABLE above. Never write a description — copy the title exactly.

=== OUTPUT FORMAT ===
Respond ONLY with a valid JSON array. No markdown, no explanation, no code fences.

Each slide object:
{{
  "slide_number": 1,
  "layout_type": "title",
  "title": "Slide Title Here",
  "key_points": ["point 1", "point 2"],
  "use_table": null,
  "metadata": {{}},
  "speaker_notes": "One sentence summary"
}}

For two_col_compare, metadata must include left_label and right_label.
For kpi_stats, key_points must follow "LABEL: VALUE: context" format.
For timeline, key_points must follow "YEAR: Event" format.
For grid_4col, key_points must follow "HEADER: description" format (exactly 4 items).
"""

    response = llm.invoke([HumanMessage(content=prompt)])

    try:
        text = response.content.strip()

        # Robust JSON extraction
        match = re.search(r'\[.*\]', text, re.DOTALL)
        if match:
            text = match.group(0)
        else:
            text = re.sub(r'^```(?:json)?\s*', '', text)
            text = re.sub(r'\s*```$', '', text)
            text = text.strip()

        slides = json.loads(text)

        # Sanitize every slide
        slides = [sanitize_slide(slide, i) for i, slide in enumerate(slides)]

        # Enforce slide 1 = title, slide 2 = executive_summary, last = conclusion
        if slides and slides[0]["layout_type"] != "title":
            slides[0]["layout_type"] = "title"
        if len(slides) > 1 and slides[1]["layout_type"] != "executive_summary":
            slides[1]["layout_type"] = "executive_summary"
        if slides and slides[-1]["layout_type"] != "conclusion":
            slides[-1]["layout_type"] = "conclusion"

        print(f"✅ LLM planned {len(slides)} slides")
        for s in slides:
            print(f"   Slide {s['slide_number']}: [{s['layout_type']}] {s['title']}")

        return slides

    except json.JSONDecodeError as e:
        print(f"❌ JSON parse error: {e}")
        print(f"Raw response preview: {response.content[:300]}")
        return get_fallback_plan(parsed_doc)


def get_fallback_plan(parsed_doc):
    """Robust fallback plan if LLM fails entirely."""
    print("⚠️  Using fallback slide plan")
    slides = []

    # Slide 1: Title
    slides.append({
        "slide_number": 1, "layout_type": "title",
        "title": parsed_doc["title"],
        "key_points": [parsed_doc.get("subtitle", "")],
        "use_table": None, "metadata": {}, "speaker_notes": "Title slide"
    })

    # Slide 2: Executive Summary
    exec_points = []
    if parsed_doc.get("executive_summary"):
        sentences = [s.strip() for s in parsed_doc["executive_summary"].split('.') if len(s.strip()) > 20]
        exec_points = [f"Key Insight {i+1}: {s[:80]}" for i, s in enumerate(sentences[:5])]
    if not exec_points:
        exec_points = [f"Key Finding {i+1}: Important insight from the document" for i in range(5)]

    slides.append({
        "slide_number": 2, "layout_type": "executive_summary",
        "title": "Executive Summary",
        "key_points": exec_points[:5],
        "use_table": None, "metadata": {}, "speaker_notes": "Executive summary"
    })

    # Content slides from sections
    main_sections = [
        s for s in parsed_doc["sections"]
        if not s.get("is_executive_summary") and not s.get("is_references")
        and "table of contents" not in s["title"].lower()
    ][:8]

    layout_cycle = ["split_panel", "split_panel", "grid_4col", "split_panel", "two_col_compare", "split_panel", "split_panel", "split_panel"]

    for i, section in enumerate(main_sections):
        layout = layout_cycle[i % len(layout_cycle)]
        slides.append({
            "slide_number": len(slides) + 1,
            "layout_type": layout,
            "title": section["title"][:60],
            "key_points": ["Key point from this section", "Supporting detail", "Additional insight"],
            "use_table": None,
            "metadata": {"left_label": "Strengths", "right_label": "Challenges"} if layout == "two_col_compare" else {},
            "speaker_notes": f"Section: {section['title']}"
        })

    # Add table + chart if available
    if parsed_doc.get("tables"):
        t = parsed_doc["tables"][0]
        slides.append({
            "slide_number": len(slides) + 1, "layout_type": "data_table",
            "title": t.get("title", "Data Overview"),
            "key_points": [], "use_table": t.get("title", ""),
            "metadata": {}, "speaker_notes": "Data table slide"
        })
        slides.append({
            "slide_number": len(slides) + 1, "layout_type": "chart",
            "title": "Data Visualization",
            "key_points": [], "use_table": t.get("title", ""),
            "metadata": {}, "speaker_notes": "Chart slide"
        })

    # Conclusion
    slides.append({
        "slide_number": len(slides) + 1, "layout_type": "conclusion",
        "title": "Key Takeaways",
        "key_points": ["Strategic insight from the analysis", "Key recommendation", "Forward looking perspective"],
        "use_table": None, "metadata": {}, "speaker_notes": "Conclusion"
    })

    return slides


if __name__ == "__main__":
    from parser import parse_markdown

    API_KEY = "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"

    doc = parse_markdown("Accenture Tech Acquisition Analysis.md")
    slides = plan_slides(doc, API_KEY)

    print("\nFinal Slide Plan:")
    for slide in slides:
        print(f"  Slide {slide['slide_number']}: [{slide['layout_type']}] {slide['title']}")
        for point in slide.get('key_points', [])[:2]:
            print(f"    • {point}")