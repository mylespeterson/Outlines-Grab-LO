# Outlines-Grab-LO

Automated tool that:
1. Reads course codes (e.g. `CET1023`, `CMP1117`, `NET2002`) from your Google Sheets spreadsheet
2. Searches the [Cambrian College course outline portal](https://cf.cambriancollege.ca/course_outlines/offerings/SearchPages/courseselect.cfm) for each course
3. Downloads the course outline PDFs
4. Extracts **Learning Outcomes** and **Objectives** from each PDF
5. Saves everything to a single Excel workbook (`course_learning_outcomes.xlsx`)

---

## Requirements

- Python 3.10 or later
- The Google Sheets spreadsheet must be shared as **"Anyone with the link can view"**

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python course_outline_fetcher.py
```

The script will:
- Fetch the spreadsheet at the configured Google Sheets URL
- Search the Cambrian website for each course prefix (term `202526`, year `all`)
- Download and parse each PDF
- Write `course_learning_outcomes.xlsx` in the current directory

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--sheets-url` | built-in URL | Google Sheets CSV export URL |
| `--output` | `course_learning_outcomes.xlsx` | Output file path |
| `--pdf-dir` | temporary directory | Directory to save downloaded PDFs |

**Keep PDFs after the run:**
```bash
python course_outline_fetcher.py --pdf-dir ./pdfs
```

**Use a different spreadsheet:**
```bash
python course_outline_fetcher.py --sheets-url "https://docs.google.com/spreadsheets/d/<ID>/export?format=csv&gid=<GID>"
```

## Output

The Excel workbook contains two sheets:

| Sheet | Contents |
|-------|----------|
| **Summary** | One row per course — code, name, number of LOs found, any errors |
| **Learning Outcomes** | One row per outcome, grouped by course |

## Running Tests

```bash
pip install pytest reportlab
python -m pytest test_course_outline_fetcher.py -v
```
