import fitz
import time
from pathlib import Path

pdf_path = Path.home() / "Desktop" / "Gazette 10th Class FIRST ANNUAL 2026.pdf"
search_name = "Tahreem Kashif"

start = time.time()

doc = fitz.open(pdf_path)

print(f"PDF pages: {len(doc):,}")
print(f'Searching for: "{search_name}"')
print("-" * 60)

found = []

for page_no, page in enumerate(doc):
    text = page.get_text("text")

    if search_name.lower() in text.lower():
        found.append(page_no + 1)

        print(f"\nFOUND on PDF page {page_no + 1}")

        # Show lines around the matching name
        lines = text.splitlines()

        for i, line in enumerate(lines):
            if search_name.lower() in line.lower():
                start_line = max(0, i - 3)
                end_line = min(len(lines), i + 4)

                print("\n".join(lines[start_line:end_line]))
                print("-" * 60)

doc.close()

elapsed = time.time() - start

print()

if found:
    print(f"Found '{search_name}' on page(s): {found}")
else:
    print(f"'{search_name}' was not found.")

print(f"Search completed in {elapsed:.2f} seconds.")