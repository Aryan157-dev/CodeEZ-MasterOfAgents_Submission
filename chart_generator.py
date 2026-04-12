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
            # Try all columns for a numeric value, not just column 1
            val = None
            for cell in row[1:]:
                # Extract first number found in cell
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
                labels.append(label[:20])  # truncate long labels
                values.append(val)

    x_label = headers[0] if headers else ""
    y_label = headers[1] if len(headers) > 1 else "Value"
    return labels, values, x_label, y_label


def _detect_chart_type(table):
    """
    Intelligently pick chart type based on table content.
    - Few rows + percentage-like values → donut
    - Many rows + time/year labels → line
    - Categorical + many items → horizontal bar
    - Default → vertical bar
    """
    headers = table.get('headers', [])
    rows    = table.get('rows', [])
    labels, values, _, _ = _parse_values(table)

    if not values:
        return "bar"

    # Check for year/time series
    first_col = [str(r[0]) for r in rows if r]
    has_years = sum(1 for l in first_col if re.match(r'20\d{2}|19\d{2}', l)) >= 2
    if has_years and len(values) >= 3:
        return "line"

    # Check for percentage data → donut
    all_pct = all(0 <= v <= 100 for v in values)
    if all_pct and len(values) <= 6:
        return "donut"

    # Many categories → horizontal bar
    if len(values) > 5:
        return "horizontal_bar"

    return "bar"


# ── Chart builders ────────────────────────────────────────────────────────
def make_bar_chart(table, output_path):
    labels, values, x_label, y_label = _parse_values(table)
    if not values:
        return None

    fig, ax = plt.subplots(figsize=(11, 5))
    fig.patch.set_facecolor('white')
    ax.set_facecolor(LIGHT)

    # Gradient-feel bars using alpha variation
    colors = [PRIMARY] * len(values)
    bars = ax.bar(labels, values, color=colors, width=0.5, zorder=3,
                  edgecolor='white', linewidth=0.8)

    # Value labels
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2,
                bar.get_height() + max(values) * 0.015,
                f'{val:g}',
                ha='center', va='bottom',
                fontsize=11, fontweight='bold', color=DARK)

    ax.set_xlabel(x_label, fontsize=12, color=DARK, labelpad=10)
    ax.set_ylabel(y_label, fontsize=12, color=DARK, labelpad=10)
    ax.tick_params(colors=DARK, labelsize=10)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#CCCCCC')
    ax.spines['bottom'].set_color('#CCCCCC')
    ax.yaxis.grid(True, color='#DDDDDD', linestyle='--', linewidth=0.7, zorder=0)
    ax.set_axisbelow(True)

    # Wrap long labels
    if labels and max(len(l) for l in labels) > 10:
        ax.set_xticklabels([l[:15] + '...' if len(l) > 15 else l for l in labels],
                           rotation=20, ha='right')

    plt.tight_layout()
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
    fill_pcts   = [0.85, 0.65, 0.75, 0.45]

    for i, (label, value, max_val, unit) in enumerate(stats_list[:n]):
        ax = axes[i]
        ax.set_facecolor(DARK)

        pct = fill_pcts[i % 4]
        ring_color = ring_colors[i % len(ring_colors)]

        # Thin background ring
        bg = mpatches.Wedge((0.5, 0.5), 0.32, 0, 360,
                             width=0.06, color='#2A2A2A')
        ax.add_patch(bg)

        # Thin progress ring
        prog = mpatches.Wedge((0.5, 0.5), 0.32, 90,
                               90 - (pct * 360),
                               width=0.06, color=ring_color)
        ax.add_patch(prog)

        # Small dot at end of progress
        end_angle = (90 - pct * 360) * (3.14159 / 180)
        dot_x = 0.5 + 0.32 * __import__('math').cos(end_angle)
        dot_y = 0.5 + 0.32 * __import__('math').sin(end_angle)
        dot = plt.Circle((dot_x, dot_y), 0.03, color=ring_color)
        ax.add_patch(dot)

        # Value — clean and centered
        ax.text(0.5, 0.58, f'{value:g}{unit}',
                ha='center', va='center',
                fontsize=14, fontweight='bold',
                color='white',
                transform=ax.transAxes)

        # Percentage — small grey below
        ax.text(0.5, 0.42, f'{int(pct * 100)}%',
                ha='center', va='center',
                fontsize=8, color=ring_color,
                transform=ax.transAxes)

        # Thin divider line
        ax.axhline(y=0.2, xmin=0.2, xmax=0.8,
                   color='#333333', linewidth=0.5)

        # Label — small caps at bottom
        ax.text(0.5, 0.1, label.upper(),
                ha='center', va='center',
                fontsize=7, color="#F18509",
                fontfamily='monospace',
                transform=ax.transAxes)

        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis('off')

    # Thin separating lines between rings
    for j in range(1, n):
        fig.add_artist(plt.Line2D(
            [j / n, j / n], [0.1, 0.9],
            color='#333333', linewidth=0.8,
            transform=fig.transFigure
        ))

    plt.subplots_adjust(wspace=0.05, left=0.02, right=0.98,
                        top=0.95, bottom=0.05)
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor=DARK, edgecolor='none')
    plt.close()
    return output_path


def make_horizontal_bar_chart(table, output_path):
    labels, values, x_label, y_label = _parse_values(table)
    if not values:
        return None

    # Sort by value descending
    paired = sorted(zip(values, labels), reverse=True)
    values = [p[0] for p in paired]
    labels = [p[1] for p in paired]

    fig, ax = plt.subplots(figsize=(11, max(4, len(values) * 0.6)))
    fig.patch.set_facecolor('white')
    ax.set_facecolor(LIGHT)

    # Color gradient — first bar darkest
    n = len(values)
    bar_colors = [PRIMARY] * n

    bars = ax.barh(labels, values, color=bar_colors, height=0.55,
                   edgecolor='white', linewidth=0.8, zorder=3)

    # Value labels
    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + max(values) * 0.01, bar.get_y() + bar.get_height() / 2,
                f'{val:g}',
                va='center', ha='left',
                fontsize=10, fontweight='bold', color=DARK)

    ax.set_xlabel(y_label, fontsize=12, color=DARK, labelpad=10)
    ax.tick_params(colors=DARK, labelsize=10)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#CCCCCC')
    ax.spines['bottom'].set_color('#CCCCCC')
    ax.xaxis.grid(True, color='#DDDDDD', linestyle='--', linewidth=0.7, zorder=0)
    ax.set_axisbelow(True)
    ax.invert_yaxis()

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    return output_path


def make_line_chart(table, output_path):
    labels, values, x_label, y_label = _parse_values(table)
    if not values or len(values) < 2:
        return make_bar_chart(table, output_path)

    fig, ax = plt.subplots(figsize=(11, 5))
    fig.patch.set_facecolor('white')
    ax.set_facecolor(LIGHT)

    x = range(len(labels))

    # Fill under line
    ax.fill_between(x, values, alpha=0.15, color=PRIMARY)

    # Line + markers
    ax.plot(x, values, color=PRIMARY, linewidth=2.5, zorder=4,
            marker='o', markersize=8, markerfacecolor=PRIMARY,
            markeredgecolor='white', markeredgewidth=2)

    # Value labels
    for xi, val in zip(x, values):
        ax.text(xi, val + max(values) * 0.02, f'{val:g}',
                ha='center', va='bottom',
                fontsize=10, fontweight='bold', color=DARK)

    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=10, color=DARK)
    ax.set_xlabel(x_label, fontsize=12, color=DARK, labelpad=10)
    ax.set_ylabel(y_label, fontsize=12, color=DARK, labelpad=10)
    ax.tick_params(colors=DARK)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#CCCCCC')
    ax.spines['bottom'].set_color('#CCCCCC')
    ax.yaxis.grid(True, color='#DDDDDD', linestyle='--', linewidth=0.7, zorder=0)
    ax.set_axisbelow(True)

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    return output_path


def make_donut_chart(table, output_path):
    labels, values, _, _ = _parse_values(table)
    if not values:
        return None

    # Limit to 6 slices max
    if len(values) > 6:
        labels  = labels[:6]
        values  = values[:6]

    # Color palette based on primary
    base   = PRIMARY
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

    # Centre label
    ax.text(0, 0, f'{len(values)}\nCategories',
            ha='center', va='center',
            fontsize=12, fontweight='bold', color=DARK)

    # Legend
    ax.legend(wedges, [f'{l}: {v:g}' for l, v in zip(labels, values)],
              loc='center left', bbox_to_anchor=(0.85, 0, 0.5, 1),
              fontsize=9, frameon=False)

    ax.set_aspect('equal')
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    return output_path


# ── Main entry point ──────────────────────────────────────────────────────
def generate_chart(parsed_doc, output_path="chart.png", table=None):
    tables = parsed_doc.get('tables', [])
    if not tables and table is None:
        return None

    if table is None:
        table = extract_chart_data(tables, "")
    if not table:
        return None

    # Always use donut chart — most reliable across different data types
    tables_to_try = [table] + [t for t in tables if t != table]
    
    for t in tables_to_try:
        result = make_donut_chart(t, output_path)
        if result:
            print(f"   ✅ Chart saved: {output_path}")
            return result
    
    print(f"   ⚠️ Could not generate chart")
    return None


# ── Test ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    from parser import parse_markdown
    doc = parse_markdown("Accenture Tech Acquisition Analysis.md")
    generate_chart(doc, "chart_test.png")
    print("Done! Check chart_test.png")