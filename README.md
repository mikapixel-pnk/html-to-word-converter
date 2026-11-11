# HTML to Word Converter

A practical **HTML → Microsoft Word (DOCX)** conversion demonstration focused on maintaining document structure—not just copying text.

**Created by Novel Hari Keswono** | IT Support · PHP/Laravel Developer · Automation

> **Portfolio demonstration:** All legal-style text and organization names in this project are fictional. This is not client work, a law-firm template, or legal advice.

## What the demo demonstrates

- Converts structured HTML into an **editable Microsoft Word document**
- Uses actual Word heading styles (Heading 1 / Heading 2)
- Supports automatic **multilevel section numbering** (e.g., 1, 1.1, 1.2)
- Supports nested clause numbering **(a), (b), (i), (ii)**
- Creates an **updateable Table of Contents (TOC)** using a Word field
- Retains key inline formatting: **bold**, *italic*, lists and editable tables
- Uses a conservative black-and-white, U.S. legal-style layout for illustration

## Project files

| File | Purpose |
| --- | --- |
| `examples/legal-agreement.html` | Fictional structured HTML source |
| `converter.py` | Python converter for this **defined demonstration HTML structure** |
| `examples/legal-agreement.docx` | Converted Word example (to be added alongside the code) |
| `requirements.txt` | Python dependencies |

## Run locally

Python 3.10+ recommended.

```bash
python -m pip install -r requirements.txt
python converter.py examples/legal-agreement.html output.docx
```

Open `output.docx` in **desktop Microsoft Word**. To generate or refresh the Table of Contents, right-click the TOC and select **Update Field → Update entire table**. Alternatively, press **Ctrl+A**, then **F9**.

**Quality checks:** add/remove a heading or list item, verify Word renumbers clauses automatically, refresh the TOC, inspect page breaks, and compare the output with the HTML source.

## Working with real client documents

Real firms and organizations have their own templates. This sample uses one neutral layout to illustrate conversion mechanics. In an actual engagement, the **client's Word template, paragraph/heading styles, numbering rules, margins, headers/footers, and original text take priority**.

This sample is **not a general-purpose HTML-to-DOCX engine**. Complex HTML/CSS, tracked changes, footnotes, cross-references, or legal-specific numbering schemes may require additional implementation and validation.

## About

**Novel Hari Keswono** — Pontianak, Indonesia

IT support, PHP/Laravel development, website maintenance, data processing and automation.

[GitHub Profile](https://github.com/mikapixel-pnk)
