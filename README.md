# Italify📝

**Italify** is a Python script that automatically detects **scientific names** in Word documents (`.docx` or `.doc`) and converts them into **italicized format**.  

It is designed to simplify **research writing**, ensuring consistent formatting of scientific names across manuscripts, reports, and articles, saving time and greatly reducing manual editing.

---

## Features
- Detects **binomial scientific names** (e.g., *Escherichia coli*, *Homo sapiens*).  
- Validates names against the **NCBI Taxonomy database** via Biopython’s Entrez API.  
- Supports both `.doc` and `.docx` files (with automatic conversion of `.doc` → `.docx`).  
- Applies italic formatting only to verified scientific names.  
- User-friendly: prompts for input and output file locations using file dialogs.  
- Saves processed files with a suffix (like, `input.file.name_itzd.docx`).  

---

## Requirements

Make sure you have **Python 3.8+** installed. Then install the dependencies:

```bash
pip install biopython python-docx pywin32
