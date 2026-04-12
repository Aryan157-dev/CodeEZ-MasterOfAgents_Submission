# 🤖 CodeEZ: Master of Agents — MD to PPTX Agentic Pipeline

**Submitted by:** Aryan Verma  
**Hackathon:** Code EZ: Master of Agents — EZ Works  
**GitHub:** https://github.com/Aryan157-dev/CodeEZ-MasterOfAgents_Submission

---

## 📌 Overview

An intelligent, agentic pipeline that converts any Markdown (`.md`) file into a professional, visually rich PowerPoint presentation (`.pptx`) — automatically.

The system uses a **multi-agent LangChain + Groq LLM architecture** to parse documents, plan slide layouts intelligently, generate programmatic charts and infographics, and build polished presentations — all from a single command.

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
      ├──────────────────────────┐
      ▼                          ▼
┌─────────────────┐    ┌──────────────────┐
│ chart_generator  │    │  pptx_builder.py  │
│      .py         │    │                   │
│                  │    │  10 slide layouts  │
│ Auto-detects     │    │  built with        │
│ best chart type: │    │  python-pptx       │
│ bar/line/donut/  │    │                   │
│ horizontal bar / │    │  Multi-theme       │
│ progress rings   │    │  support           │
└─────┬───────────┘    └─────┬─────────────┘
      │                       │
      └──────────┬────────────┘
                 ▼
         Output (.pptx file)
```

---

## 🧠 Intelligent Layout System

The LLM agent selects from **10 distinct slide layouts** based on content analysis:

| Layout | When Used |
|--------|-----------|
| `title` | Opening slide |
| `executive_summary` | 5-card summary with header + description |
| `split_panel` | General content — bold left panel + bullet cards |
| `timeline` | Historical progression, milestones, year-by-year |
| `kpi_stats` | Large bold numbers — financial metrics |
| `kpi_visual` | Dark dashboard with circular progress rings |
| `two_col_compare` | Successes vs Challenges, A vs B comparisons |
| `grid_4col` | 4-pillar frameworks, strategic domains |
| `data_table` | Tabular data from markdown tables |
| `chart` | Auto-generated bar/line/donut/horizontal bar charts |
| `conclusion` | Key takeaways with arrow bullets |

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

### Installation

```bash
# Clone the repository
git clone https://github.com/Aryan157-dev/CodeEZ-MasterOfAgents_Submission.git
cd CodeEZ-MasterOfAgents_Submission

# Install dependencies
pip install python-pptx langchain-groq langchain-core matplotlib python-dotenv numpy
```

### API Key Setup

Create a `.env` file in the project root:

```
GROQ_API_KEY=your_groq_api_key_here
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
├── pptx_builder.py      # PPTX builder — 10 layout builders + theme support
├── chart_generator.py   # Chart generator — 4 chart types + progress rings
├── .env                 # API key (not pushed to GitHub)
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
Instead of mapping every section to the same bullet-point layout, the LLM agent analyzes content signals — presence of years (→ timeline), percentages (→ donut chart), comparisons (→ two_col_compare) — and selects the most appropriate layout per slide. This produces presentations that feel designed, not generated.

### 2. Modular Architecture
Each component has a single responsibility:
- `parser.py` — content extraction only
- `slide_planner.py` — layout planning only  
- `pptx_builder.py` — rendering only
- `chart_generator.py` — visualization only

This makes the system easy to extend, debug, and maintain.

### 3. Smart Chart Type Detection
The chart generator auto-detects the best visualization:
- Year-based data → Line chart
- Percentage data ≤6 items → Donut chart
- Many categories → Horizontal bar chart
- Default → Vertical bar chart

### 4. Graceful Fallbacks
Every component has fallback behavior:
- LLM JSON parse failure → rule-based fallback plan
- No numerical data → skip chart gracefully
- Unknown layout type → defaults to split_panel
- Missing table → uses first available table

### 5. Theme System
A centralized theme dictionary controls all colors across every layout and chart. Switching themes with a single CLI argument changes the entire visual identity of the presentation.

---

## 📊 Sample Outputs

See the `TestCases/` folder for sample markdown inputs and their generated PPTX outputs across different domains:
- Corporate strategy analysis
- Financial benchmarking
- Market research reports
- Technology sector analysis

---

## 🛠 Dependencies

```
python-pptx
langchain-groq
langchain-core
matplotlib
numpy
python-dotenv
```

Install all with:
```bash
pip install python-pptx langchain-groq langchain-core matplotlib python-dotenv numpy
```

---

## 👨‍💻 Author

**Aryan Verma**  
BTech Computer Science — Sharda University  
GitHub: [@Aryan157-dev](https://github.com/Aryan157-dev)
