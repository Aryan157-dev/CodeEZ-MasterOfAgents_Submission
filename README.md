# 🤖 CodeEZ: Master of Agents — MD to PPTX Agentic Pipeline

**Submitted by:** Aryan Verma  
**Hackathon:** Code EZ: Master of Agents — EZ Works  
**GitHub:** https://github.com/Aryan157-dev/CodeEZ-MasterOfAgents_Submission

---

## 📌 Overview

An intelligent, agentic pipeline that converts any Markdown (`.md`) file into a professional, visually rich PowerPoint presentation (`.pptx`) — automatically.

The system uses a **multi-agent LangChain + Groq LLM architecture** to parse documents, plan slide layouts intelligently, fetch contextually relevant images, generate programmatic charts and infographics, and build polished presentations — all from a single command.

---

## 🏗 System Architecture

```
Input (.md file)
       │
       ▼
┌─────────────┐
│   parser.py  │  ← Extracts sections, tables, key stats,
│              │    timeline events, comparison signals
└─────┬───────┘
      │
      ▼
┌──────────────────┐
│ slide_planner.py  │  ← LangChain + Groq LLM agent
│                   │    Plans 10-13 slides with intelligent
│                   │    layout selection per slide
└─────┬────────────┘
      │
      ├──────────────────────────┬─────────────────────┐
      ▼                          ▼                     ▼
┌─────────────────┐    ┌──────────────────┐   ┌─────────────────┐
│ chart_generator  │    │  pptx_builder.py  │   │ image_fetcher.py │
│      .py         │    │                   │   │                  │
│                  │    │  12 slide layouts  │   │ Pexels API       │
│ Smart chart type │    │  built with        │   │ Fetches relevant │
│ detection:       │    │  python-pptx       │   │ images per slide │
│ bar/line/donut/  │    │                    │   │ based on topic + │
│ horizontal bar / │    │  Multi-theme        │   │ slide keyword    │
│ progress rings   │    │  support            │   │                  │
└─────┬───────────┘    └─────┬─────────────┘   └────────┬────────┘
      │                       │                           │
      └──────────┬────────────┘───────────────────────────┘
                 ▼
         Output (.pptx file)
```

---

## 🧠 Intelligent Layout System

The LLM agent selects from **12 distinct slide layouts** based on content analysis:

| Layout | When Used |
|--------|-----------|
| `title` | Opening slide with full-bleed topic image |
| `executive_summary` | 5-card summary with header + description |
| `split_panel` | General content — bold left panel with contextual image + bullet cards |
| `timeline` | Historical progression — alternating cards above/below spine |
| `kpi_stats` | Large bold numbers — financial metrics |
| `kpi_visual` | Dark dashboard with circular progress rings |
| `two_col_compare` | Successes vs Challenges, A vs B comparisons |
| `grid_4col` | 4-pillar frameworks, strategic domains |
| `data_table` | Tabular data with dynamic row sizing and insight bar |
| `chart` | Auto-generated bar/line/donut/horizontal bar charts |
| `big_statement` | Single landmark stat or insight — full-slide bold statement |
| `icon_row` | Numbered pillar cards for processes and frameworks |
| `conclusion` | Key takeaways with arrow bullets + background image |

---

## 🖼 Contextual Image Integration

Every slide now fetches a relevant image automatically via the **Pexels API**:

- **Title slides** — full right-half image based on document topic
- **Split panel slides** — image in left panel with semi-transparent dark overlay for readability
- **Conclusion slides** — full background image
- **Timeline slides** — clean white background (no image, for clarity)
- **Chart/table slides** — white background (images would distract from data)

The keyword for each image is extracted by combining the **slide title** with the **document title** — so "Market Overview" in a taxi report searches "market overview taxi" not just "market".

---

## 📊 Smart Chart System

Charts are now intelligently typed and professionally styled:

| Data Pattern | Chart Type |
|---|---|
| Year-based labels (2020, 2021...) with 3+ points | Line chart |
| Values that sum to ~100% | Donut chart |
| More than 5 categories | Horizontal bar chart |
| Default | Vertical bar chart |

All charts feature clean white backgrounds, minimal gridlines, bold value labels, and proper titles. Chart sizes are compact to sit proportionally within the slide panel.

---

## 🎨 Multi-Theme Support

Three built-in color themes:

| Theme | Colors | Best For |
|-------|--------|----------|
| `red` | Red + Dark Charcoal | Corporate, Strategy |
| `green` | Forest Green + Dark Green | Sustainability, Finance |
| `blue` | Navy Blue + Dark Navy | Technology, Healthcare |

---

## ⚙️ Setup Instructions

### Prerequisites
- Python 3.10+
- A free Groq API key from [console.groq.com](https://console.groq.com)
- A free Pexels API key from [pexels.com/api](https://www.pexels.com/api/)

### Installation

```bash
# Clone the repository
git clone https://github.com/Aryan157-dev/CodeEZ-MasterOfAgents_Submission.git
cd CodeEZ-MasterOfAgents_Submission

# Install dependencies
pip install python-pptx langchain-groq langchain-core matplotlib python-dotenv numpy requests
```

### API Key Setup

Create a `.env` file in the project root:

```
GROQ_API_KEY=your_groq_api_key_here
PEXELS_API_KEY=your_pexels_api_key_here
```

---

## 🚀 How to Run

### Basic Usage

```bash
python main.py "path/to/your/file.md"
```

### With Theme Selection

```bash
python main.py "path/to/your/file.md" green
python main.py "path/to/your/file.md" red
python main.py "path/to/your/file.md" blue
```

### Examples

```bash
python main.py "TestCases/Accenture Tech Acquisition Analysis.md" red
python main.py "TestCases/NYSE Stock Valuation Multiples Analysis.md" green
python main.py "TestCases/Used Commercial Taxi Market in India.md" blue
```

The output `.pptx` is saved in the same folder as the input `.md` file.

---

## 📁 Project Structure

```
CodeEZ-MasterOfAgents_Submission/
│
├── main.py              # Entry point — CLI interface
├── parser.py            # Markdown parser — extracts content, tables, stats
├── slide_planner.py     # LangChain + Groq LLM agent — plans slide layouts
├── pptx_builder.py      # PPTX builder — 12 layout builders + theme support
├── chart_generator.py   # Chart generator — smart chart types + progress rings
├── image_fetcher.py     # Image fetcher — Pexels API integration
├── .env                 # API keys (not pushed to GitHub)
├── .gitignore
│
└── TestCases/           # Sample inputs and generated outputs
    ├── Accenture Tech Acquisition Analysis.md
    ├── Accenture Tech Acquisition Analysis.pptx
    ├── NYSE Stock Valuation Multiples Analysis.md
    ├── NYSE Stock Valuation Multiples Analysis.pptx
    └── ... (more MD + PPTX pairs)
```

---

## 🔑 Key Design Decisions

### 1. Agentic Layout Intelligence
Instead of mapping every section to the same bullet-point layout, the LLM agent analyzes content signals — presence of years (→ timeline), percentages (→ donut chart), comparisons (→ two_col_compare), landmark stats (→ big_statement), process steps (→ icon_row) — and selects the most appropriate layout per slide. This produces presentations that feel designed, not generated.

### 2. Modular Architecture
Each component has a single responsibility:
- `parser.py` — content extraction only
- `slide_planner.py` — layout planning only
- `pptx_builder.py` — rendering only
- `chart_generator.py` — visualization only
- `image_fetcher.py` — image retrieval only

### 3. Smart Chart Type Detection
The chart generator auto-detects the best visualization based on data patterns:
- Year-based labels → Line chart
- Values summing to ~100% → Donut chart
- More than 5 categories → Horizontal bar chart
- Default → Vertical bar chart

### 4. Contextual Image Fetching
`image_fetcher.py` combines the slide title keyword with the document title to generate context-aware Pexels queries. Images are cached in-memory to avoid redundant API calls. Backgrounds are applied selectively — only on slides where they enhance rather than distract.

### 5. Rich LLM Content Prompting
The slide planner prompt enforces **20-30 word bullet points** with specific facts, numbers, percentages, and named examples from the document. BAD/GOOD examples are embedded in the prompt to guide the LLM away from generic summaries.

### 6. Graceful Fallbacks
Every component has fallback behavior:
- LLM JSON parse failure → rule-based fallback plan
- No numerical data → skip chart gracefully
- Pexels API unavailable → falls back to plain colored background
- Unknown layout type → defaults to `split_panel`
- Missing table → uses first available table

### 7. Theme System
A centralized theme dictionary controls all colors across every layout and chart. Switching themes with a single CLI argument changes the entire visual identity of the presentation.

---

## 📊 Sample Outputs

See the `TestCases/` folder for sample markdown inputs and their generated PPTX outputs across different domains:
- Corporate strategy analysis (Accenture Tech Acquisitions)
- Financial benchmarking (NYSE Stock Valuation)
- Market research reports (Used Commercial Taxi Market India)
- Technology sector analysis (AI Bubble Detection)

---

## 🛠 Dependencies

```
python-pptx
langchain-groq
langchain-core
matplotlib
numpy
python-dotenv
requests
```

Install all with:
```bash
pip install python-pptx langchain-groq langchain-core matplotlib python-dotenv numpy requests
```

---

## 👨‍💻 Author

**Aryan Verma**  
BTech Computer Science — Sharda University  
GitHub: [@Aryan157-dev](https://github.com/Aryan157-dev)