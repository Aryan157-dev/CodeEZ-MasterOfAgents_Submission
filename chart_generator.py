import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os
import re

# ── Theme colors (updated by set_chart_theme) ────────────────────────────
PRIMARY = '#EF4444'
DARK    = '#2C2C2C'
LIGHT   = '#F5F5F5'
MID     = '#A0A0A0'

CHART_THEMES = {
    "red":   {"primary": "#EF4444", "secondary": "#FF8080", "dark": "#2C2C2C"},
    "green": {"primary": "#4A7C59", "secondary": "#7DB892", "dark": "#1E3A2F"},
    "blue":  {"primary": "#1A6BB5", "secondary": "#5A9FD4", "dark": "#0D2B4E"},
}

def set_chart_theme(theme_name="green"):
    global PRIMARY, DARK
    t = CHART_THEMES.get(theme_name, CHART_THEMES["green"])
    PRIMARY = t["primary"]
    DARK    = t["dark"]


# ── Data extraction ───────────────────────────────────────────────────────
def extract_chart_data(tables, sections_content=""):
    """Find the best table to chart — prefers tables with numerical data."""
    for table in tables:
        rows = table.get('rows', [])
        if len(rows) < 2:
            continue
        for row in rows:
            for cell in row[1:]:
                clean = re.sub(r'[,$%\s]', '', str(cell))
                if re.match(r'^\d+\.?\d*$', clean):
                    return table
    return tables[0] if tables else None


def _parse_values(table):
    headers = table.get('headers', [])
    rows    = table.get('rows', [])
    labels, values = [], []

    for row in rows:
        if len(row) >= 2:
            label = str(row[0]).strip()
            val = None
            for cell in row[1:]:
                nums = re.findall(r'[\d,]+\.?\d*', str(cell).replace(',', ''))
                for n in nums:
                    try:
                        v = float(n)
                        if v > 0:
                            val = v
                            break
                    except ValueError:
                        continue
                if val is not None:
                    break
            if val is not None:
                labels.append(label[:20])
                values.append(val)

    x_label = headers[0] if headers else ""
    y_label = headers[1] if len(headers) > 1 else "Value"
    return labels, values, x_label, y_label


def _detect_chart_type(table):
    """
    Intelligently pick chart type based on table content:
    - Year/time labels with 3+ rows → line
    - Few rows + percentage-like values → donut
    - Many rows (>5) → horizontal bar
    - Default → bar
    """
    rows   = table.get('rows', [])
    labels, values, _, _ = _parse_values(table)

    if not values:
        return "bar"

    # Time series detection
    first_col = [str(r[0]) for r in rows if r]
    has_years = sum(1 for l in first_col if re.match(r'20\d{2}|19\d{2}', l)) >= 2
    if has_years and len(values) >= 3:
        return "line"

    # Percentage / proportion data → donut
    # Only use donut if values sum close to 100 (actual proportions)
    # NOT just because values happen to be between 0-100 (e.g. P/E ratios)
    total = sum(values)
    all_pct = (85 <= total <= 115) and all(0 < v <= 100 for v in values)
    if all_pct and len(values) <= 6:
        return "donut"

    # Many categories → horizontal bar (easier to read)
    if len(values) > 5:
        return "horizontal_bar"

    return "bar"


# ── Chart builders ────────────────────────────────────────────────────────
def _make_chart_title(ax, table):
    """Add table title as chart title if available."""
    title = table.get('title', '')
    if title:
        ax.set_title(title, fontsize=13, fontweight='bold', color=DARK,
                     pad=12, loc='left')


def _style_axes(ax, table):
    """Apply consistent sharp professional styling to any axis."""
    title = table.get('title', '')
    if title:
        ax.set_title(title[:65], fontsize=13, fontweight='bold',
                     color=DARK, pad=14, loc='left')
    for spine in ['top', 'right', 'left']:
        ax.spines[spine].set_visible(False)
    ax.spines['bottom'].set_color('#E0E0E0')
    ax.yaxis.grid(True, color='#F0F0F0', linestyle='-', linewidth=1.0, zorder=0)
    ax.set_axisbelow(True)
    ax.tick_params(axis='both', length=0, labelsize=10, colors='#555555')


def make_bar_chart(table, output_path):
    labels, values, x_label, y_label = _parse_values(table)
    if not values:
        return None

    fig, ax = plt.subplots(figsize=(10, 4.2))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    x = range(len(labels))
    bars = ax.bar(x, values, color=PRIMARY, width=0.52, zorder=3,
                  edgecolor='white', linewidth=1.2)

    # Gradient alpha
    for j, bar in enumerate(bars):
        bar.set_alpha(1.0 - j * (0.5 / max(len(bars), 1)))

    # Value labels
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2,
                bar.get_height() + max(values) * 0.012,
                f'{val:g}', ha='center', va='bottom',
                fontsize=11, fontweight='bold', color=DARK)

    clean = [l[:16] + '…' if len(l) > 16 else l for l in labels]
    ax.set_xticks(list(x))
    needs_rotation = max((len(l) for l in clean), default=0) > 10
    ax.set_xticklabels(clean,
                       rotation=28 if needs_rotation else 0,
                       ha='right' if needs_rotation else 'center',
                       fontsize=10)
    ax.set_ylabel(y_label, fontsize=10, color='#777777', labelpad=6)
    _style_axes(ax, table)

    plt.tight_layout(pad=1.5)
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    return output_path


def make_progress_rings(stats_list, output_path):
    if not stats_list:
        return None

    n = min(len(stats_list), 4)
    fig, axes = plt.subplots(1, n, figsize=(11, 3.5))
    fig.patch.set_facecolor(DARK)

    if n == 1:
        axes = [axes]

    ring_colors = [PRIMARY, '#CCCCCC', '#888888', '#555555']

    for i, (label, value, max_val, unit) in enumerate(stats_list[:n]):
        ax = axes[i]
        ax.set_facecolor(DARK)

        # Compute real fill percentage from actual data
        pct = min(value / max_val, 1.0) if max_val > 0 else 0.5
        ring_color = ring_colors[i % len(ring_colors)]

        bg = mpatches.Wedge((0.5, 0.5), 0.32, 0, 360,
                             width=0.06, color='#2A2A2A')
        ax.add_patch(bg)

        prog = mpatches.Wedge((0.5, 0.5), 0.32, 90,
                               90 - (pct * 360),
                               width=0.06, color=ring_color)
        ax.add_patch(prog)

        end_angle = (90 - pct * 360) * (3.14159 / 180)
        import math
        dot_x = 0.5 + 0.32 * math.cos(end_angle)
        dot_y = 0.5 + 0.32 * math.sin(end_angle)
        dot = plt.Circle((dot_x, dot_y), 0.03, color=ring_color)
        ax.add_patch(dot)

        ax.text(0.5, 0.58, f'{value:g}{unit}',
                ha='center', va='center',
                fontsize=14, fontweight='bold',
                color='white',
                transform=ax.transAxes)

        ax.text(0.5, 0.42, f'{int(pct * 100)}%',
                ha='center', va='center',
                fontsize=8, color=ring_color,
                transform=ax.transAxes)

        ax.axhline(y=0.2, xmin=0.2, xmax=0.8,
                   color='#333333', linewidth=0.5)

        ax.text(0.5, 0.1, label.upper(),
                ha='center', va='center',
                fontsize=7, color="#F18509",
                fontfamily='monospace',
                transform=ax.transAxes)

        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis('off')

    for j in range(1, n):
        fig.add_artist(plt.Line2D(
            [j / n, j / n], [0.1, 0.9],
            color='#333333', linewidth=0.8,
            transform=fig.transFigure
        ))

    plt.subplots_adjust(wspace=0.05, left=0.02, right=0.98, top=0.95, bottom=0.05)
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor=DARK, edgecolor='none')
    plt.close()
    return output_path


def make_horizontal_bar_chart(table, output_path):
    labels, values, x_label, y_label = _parse_values(table)
    if not values:
        return None

    paired = sorted(zip(values, labels), reverse=True)
    values = [p[0] for p in paired]
    labels = [p[1] for p in paired]

    fig, ax = plt.subplots(figsize=(10, max(4.0, len(values) * 0.65)))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    n = len(values)
    bars = ax.barh(labels, values, color=PRIMARY, height=0.52,
                   edgecolor='white', linewidth=1.2, zorder=3)
    for i, bar in enumerate(bars):
        bar.set_alpha(1.0 - i * (0.4 / max(n, 1)))

    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + max(values) * 0.01,
                bar.get_y() + bar.get_height() / 2,
                f'{val:g}', va='center', ha='left',
                fontsize=10, fontweight='bold', color=DARK)

    ax.set_xlabel(y_label, fontsize=10, color='#777777', labelpad=6)
    ax.tick_params(axis='y', labelsize=10)
    ax.invert_yaxis()
    _style_axes(ax, table)
    ax.xaxis.grid(True, color='#F0F0F0', linestyle='-', linewidth=1.0, zorder=0)
    ax.yaxis.grid(False)

    plt.tight_layout(pad=1.5)
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    return output_path
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    return output_path


def make_line_chart(table, output_path):
    labels, values, x_label, y_label = _parse_values(table)
    if not values or len(values) < 2:
        return make_bar_chart(table, output_path)

    fig, ax = plt.subplots(figsize=(10, 4.2))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    x = range(len(labels))

    # Area fill with gradient
    ax.fill_between(x, values, alpha=0.12, color=PRIMARY)

    # Line
    ax.plot(x, values, color=PRIMARY, linewidth=2.8, zorder=4,
            marker='o', markersize=9,
            markerfacecolor='white',
            markeredgecolor=PRIMARY, markeredgewidth=2.5)

    # Value labels
    for xi, val in zip(x, values):
        ax.text(xi, val + max(values) * 0.025, f'{val:g}',
                ha='center', va='bottom',
                fontsize=11, fontweight='bold', color=DARK)

    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_ylabel(y_label, fontsize=10, color='#777777', labelpad=6)
    _style_axes(ax, table)

    plt.tight_layout(pad=1.5)
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    return output_path


def make_donut_chart(table, output_path):
    labels, values, _, _ = _parse_values(table)
    if not values:
        return None

    if len(values) > 6:
        labels = labels[:6]
        values = values[:6]

    shades = [PRIMARY, '#888888', '#AAAAAA', '#CCCCCC', '#555555', '#333333']

    fig, ax = plt.subplots(figsize=(9, 5))
    fig.patch.set_facecolor('white')

    wedges, texts, autotexts = ax.pie(
        values,
        labels=None,
        autopct='%1.1f%%',
        colors=shades[:len(values)],
        startangle=90,
        pctdistance=0.75,
        wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2)
    )

    for autotext in autotexts:
        autotext.set_fontsize(10)
        autotext.set_fontweight('bold')
        autotext.set_color('white')

    title = table.get('title', '')
    ax.text(0, 0, title[:12] if title else f'{len(values)}\nCategories',
            ha='center', va='center',
            fontsize=11, fontweight='bold', color=DARK)

    ax.legend(wedges, [f'{l}: {v:g}' for l, v in zip(labels, values)],
              loc='center left', bbox_to_anchor=(0.85, 0, 0.5, 1),
              fontsize=9, frameon=False)

    ax.set_aspect('equal')
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    return output_path


# ── Chart dispatch table ──────────────────────────────────────────────────
CHART_BUILDERS = {
    "bar":            make_bar_chart,
    "horizontal_bar": make_horizontal_bar_chart,
    "line":           make_line_chart,
    "donut":          make_donut_chart,
}


# ── Main entry point ──────────────────────────────────────────────────────
def generate_chart(parsed_doc, output_path="chart.png", table=None):
    tables = parsed_doc.get('tables', [])
    if not tables and table is None:
        return None

    if table is None:
        table = extract_chart_data(tables, "")
    if not table:
        return None

    chart_type = _detect_chart_type(table)
    print(f"   📊 Chart type detected: {chart_type} for table '{table.get('title', '')[:40]}'")

    builder = CHART_BUILDERS.get(chart_type, make_bar_chart)
    result  = builder(table, output_path)

    if result:
        print(f"   ✅ Chart saved: {output_path}")
    else:
        # Fallback: try bar chart
        print(f"   ⚠️  {chart_type} failed, falling back to bar chart")
        result = make_bar_chart(table, output_path)
        if result:
            print(f"   ✅ Fallback chart saved: {output_path}")
        else:
            print(f"   ❌ Could not generate chart")

    return result


# ── Test ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    from parser import parse_markdown
    doc = parse_markdown("Accenture Tech Acquisition Analysis.md")
    generate_chart(doc, "chart_test.png")
    print("Done! Check chart_test.png")