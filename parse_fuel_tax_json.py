import json
import pandas as pd
import re

# Load the JSON file
with open("beta.novascotia.ca_programs-and-services_fuel-tax-program.json", "r", encoding="utf-8") as f:
    data = json.load(f)

markdown = data.get("markdown", "")
lines = markdown.splitlines()

link_pattern = re.compile(r"\[([^\]]+)\]\((https?://[^\)]+)\)")
entries = []
breadcrumb = ""

for i, line in enumerate(lines):
    line = line.strip()

    # Update breadcrumb if we see a "##" heading
    if line.startswith("## "):
        breadcrumb = line.replace("##", "").strip()

    # Search for a link in the current line
    match = link_pattern.search(line)
    if match:
        title = match.group(1).strip()
        url = match.group(2).strip()

        # Check if next line is a valid description (not another header or empty)
        description = ""
        if i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if next_line and not next_line.startswith("#"):
                description = next_line

        # If no good description, use the title
        if not description:
            description = title

        entries.append({
            "Breadcrumb": breadcrumb,
            "Description": description,
            "URL": url
        })

# Save to Excel
df = pd.DataFrame(entries)
df.to_excel("fuel_tax_knowledge_base.xlsx", index=False)

print("✅ Excel file 'fuel_tax_knowledge_base.xlsx' created successfully.")
