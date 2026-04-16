import os
import subprocess
import sys
import time

# Theme assignment — manually tuned for each topic
THEME_MAP = {
    "accenture":    "red",
    "nyse":         "blue",
    "banking":      "blue",
    "uae":          "green",
    "solar":        "green",
    "ai":           "blue",
    "artificial":   "blue",
    "india":        "green",
    "taxi":         "green",
    "market":       "green",
    "climate":      "green",
    "energy":       "green",
    "financial":    "blue",
    "finance":      "blue",
    "stock":        "blue",
    "investment":   "blue",
    "tech":         "red",
    "digital":      "red",
    "strategy":     "red",
    "corporate":    "red",
}

def pick_theme(filename):
    """Pick theme based on filename keywords."""
    name_lower = filename.lower()
    for keyword, theme in THEME_MAP.items():
        if keyword in name_lower:
            return theme
    return "green"  # default


def run_all(test_dir="TestCases"):
    md_files = sorted([
        f for f in os.listdir(test_dir)
        if f.endswith(".md")
    ])

    if not md_files:
        print(f"❌ No .md files found in '{test_dir}' folder.")
        sys.exit(1)

    print(f"🚀 Found {len(md_files)} markdown files\n")
    print("=" * 60)

    success = []
    failed  = []

    for i, md in enumerate(md_files):
        path  = os.path.join(test_dir, md)
        theme = pick_theme(md)

        print(f"\n[{i+1}/{len(md_files)}] {md}")
        print(f"   Theme: {theme}")

        start = time.time()
        result = subprocess.run(
            ["python", "main.py", path, theme],
            capture_output=False
        )
        elapsed = round(time.time() - start, 1)

        if result.returncode == 0:
            print(f"   ✅ Done in {elapsed}s")
            success.append(md)
        else:
            print(f"   ❌ Failed after {elapsed}s")
            failed.append(md)

    print("\n" + "=" * 60)
    print(f"\n📊 Results: {len(success)} succeeded, {len(failed)} failed")

    if failed:
        print("\n❌ Failed files:")
        for f in failed:
            print(f"   - {f}")

    print("\n✅ All done! Check the TestCases folder for your .pptx files.")
    print("   Pick your best 5-6 for submission.")


if __name__ == "__main__":
    test_dir = sys.argv[1] if len(sys.argv) > 1 else "TestCases"
    run_all(test_dir)