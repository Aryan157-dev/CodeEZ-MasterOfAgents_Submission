import os
from dotenv import load_dotenv
load_dotenv()
import sys
from parser import parse_markdown
from slide_planner import plan_slides
import pptx_builder
import chart_generator
from pptx_builder import build_presentation

API_KEY = os.environ.get("GROQ_API_KEY")

def main(md_file, theme="red"):
    if not os.path.exists(md_file):
        print(f"❌ File not found: {md_file}")
        return

    # Set theme before building
    pptx_builder.set_theme(theme)
    chart_generator.set_chart_theme(theme) 
    print(f"🎨 Theme: {theme}")

    # Output file same name as input but .pptx
    output = os.path.join(os.path.dirname(md_file), 
         os.path.splitext(os.path.basename(md_file))[0] + ".pptx")

    print(f"📄 Input:  {md_file}")
    print(f"📊 Output: {output}")
    print()

    print("Step 1: Parsing markdown...")
    doc = parse_markdown(md_file)
    print(f"  ✅ Found {len(doc['sections'])} sections, {len(doc['tables'])} tables")

    print("Step 2: Planning slides with AI...")
    plan = plan_slides(doc, API_KEY)
    print(f"  ✅ Planned {len(plan)} slides")

    print("Step 3: Building presentation...")
    build_presentation(plan, doc, None, output)
    print(f"\n🎉 Done! Open {output} to view your presentation.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python main.py <markdown_file> [theme]")
        print("Themes: red, green, blue")
        print()
        print("Examples:")
        print("  python main.py 'Accenture Tech Acquisition Analysis.md'")
        print("  python main.py 'Accenture Tech Acquisition Analysis.md' red")
        print("  python main.py 'UAE Progress toward 2050 Solar Energy Targets.md' green")
        print("  python main.py 'Banking ROE Competitive Benchmarking Analysis.md' blue")
    else:
        md   = sys.argv[1]
        theme = sys.argv[2] if len(sys.argv) > 2 else "green"
        main(md, theme)