# SOW01 — D365 F&O Statement of Work Reviewer

Automated Scope of Work review tool for Dynamics 365 Finance & Operations managed services consulting. Extracts text from `.docx` SOW documents, sends it to Claude for structured assessment, and exports the results.

## Prerequisites

- Python 3.11+
- An Anthropic API key

## Installation

```bash
cd sow01
pip install -r requirements.txt
cp .env.example .env
# Edit .env and add your Anthropic API key
```

## Usage

```bash
# Stdout only (default)
python main.py --input path/to/sow.docx

# Export as markdown
python main.py --input path/to/sow.docx --format md

# Export as Word document
python main.py --input path/to/sow.docx --format docx

# Export both formats
python main.py --input path/to/sow.docx --format both

# Verbose mode (prints extracted text before API call)
python main.py --input path/to/sow.docx --format both --verbose

# Custom output directory
python main.py --input path/to/sow.docx --format md --output ./reports/
```

### Arguments

| Argument    | Required | Default    | Description                                      |
|-------------|----------|------------|--------------------------------------------------|
| `--input`   | Yes      | —          | Path to the `.docx` SOW file                     |
| `--format`  | No       | `none`     | Output format: `md`, `docx`, `both`, or `none`   |
| `--output`  | No       | `./output/`| Output directory for exported files               |
| `--verbose` | No       | `false`    | Print extracted text and diagnostics              |

### Output

Output files follow the naming pattern:

```
SOW01_[original_filename]_[YYYYMMDD].md
SOW01_[original_filename]_[YYYYMMDD].docx
```

## Project Structure

```
sow01/
├── .env.example      # API key template
├── .gitignore
├── requirements.txt
├── README.md
├── main.py           # CLI entry point
├── extractor.py      # DOCX text extraction
├── analyser.py       # Anthropic API integration
├── exporter.py       # Markdown and DOCX output
└── output/           # Default output directory
```
