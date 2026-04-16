import json
from docx import Document

# Load structured JSON
with open("QBR_Q1_FY26_Lending_Approval_to_Funding.json", "r") as f:
    data = json.load(f)

# Load template
doc = Document("templates/202604_Domain-Adapted_QBRmemo_Template vSHARE.docx")

def set_cell_text(cell, text):
    # Clear existing cell content
    cell.text = ""
    paragraphs = text.split("\n")
    cell.paragraphs[0].text = paragraphs[0] if paragraphs else ""
    for line in paragraphs[1:]:
        cell.add_paragraph(line)

def replace_in_paragraph(paragraph, old, new):
    if old in paragraph.text:
        paragraph.text = paragraph.text.replace(old, new)

def replace_everywhere(doc, old, new):
    # Normal paragraphs
    for p in doc.paragraphs:
        replace_in_paragraph(p, old, new)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, old, new)

    # Headers and footers
    for section in doc.sections:
        for p in section.header.paragraphs:
            replace_in_paragraph(p, old, new)
        for p in section.footer.paragraphs:
            replace_in_paragraph(p, old, new)

# Header fields
h = data["header"]

quarter_label = f"{h['quarter']} {h['fiscal_year']}"  # Q3 FY26

replace_everywhere(doc, "Qx FYxx", quarter_label)
replace_everywhere(doc, "FYxx", h["fiscal_year"])
replace_everywhere(doc, "[Domain]", h["domain"])
replace_everywhere(doc, "[Month] xx, xxxx", h["date"])
replace_everywhere(doc, "xxxxxx xxxxxx", h["domain_lead"])
replace_everywhere(doc, "Insert Domain Mission", h["mission"])
replace_everywhere(doc, "[Domain name]: Insert Domain Mission", h["mission"])

# Executive summary
es = data["executive_summary"]

learning_text = "\n".join([f"- {lp}" for lp in es["learning_points"]])
opportunity_text = "\n".join([f"- {op}" for op in es["next_opportunities"]])

summary_intro = f"The key priority for last quarter was to focus on {es['priority_last_quarter']}."

# Find executive summary table rows by their left-hand labels
for table in doc.tables:
    for row in table.rows:
        label = row.cells[0].text.strip()

        if label.startswith("The key priority for last quarter"):
            set_cell_text(row.cells[1], summary_intro)

        elif label.startswith("Key learnings this quarter were"):
            set_cell_text(row.cells[1], learning_text)

        elif label.startswith("Our next opportunities are"):
            set_cell_text(row.cells[1], opportunity_text)

# Save output
doc.save("QBR_Output.docx")
print("Word document generated: QBR_Output.docx")
