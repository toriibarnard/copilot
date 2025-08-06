import json
import pandas as pd
import re

# Load JSON file
with open("beta.novascotia.ca_programs-and-services_fuel-tax-program.json", "r", encoding="utf-8") as f:
    data = json.load(f)

markdown = data.get("markdown", "")
lines = markdown.splitlines()

link_pattern = re.compile(r"\[([^\]]+)\]\((https?://[^\)]+)\)")

entries = []
h1 = ""  # Main header #
h2 = ""  # Section header ##
h3 = ""  # Item header ###

for i, line in enumerate(lines):
    line = line.strip()

    if line.startswith("# "):
        h1 = line.replace("# ", "").strip()
    elif line.startswith("## "):
        h2 = line.replace("## ", "").strip()
    elif line.startswith("### "):
        h3 = line.replace("### ", "").strip()

    # Extract link
    match = link_pattern.search(line)
    if match:
        title = match.group(1).strip()
        url = match.group(2).strip()

        # Try getting description from next line
        description = ""
        if i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if next_line and not next_line.startswith("#"):
                description = next_line

        if not description:
            description = title

        # Build breadcrumb path
        breadcrumb = " > ".join(filter(None, [h1, h2, h3]))

        entries.append({
            "Breadcrumb": breadcrumb,
            "Description": description,
            "URL": url
        })

# Create and export to Excel
df = pd.DataFrame(entries)
df.to_excel("fuel_tax_breadcrumb_paths.xlsx", index=False)

print("✅ Excel with breadcrumb paths saved as 'fuel_tax_breadcrumb_paths.xlsx'")
