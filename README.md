# CrewAI DOCX Translator (Format-Preserving)

Translate full academic DOCX papers to another language while preserving the original layout (paragraph styles, tables, images). Built with CrewAI + LiteLLM (Gemini).

## Features
- Structure-preserving: edits text in place (paragraphs and table cells), keeps tables/images/styles.
- Batch translation to control token usage.
- Config-driven (language/model/paths).
- Fallback filename if the output DOCX is open.

## Project Structure
- `run.py` — main pipeline (extract → batch translate → replace → save)
- `config.yaml` — target language, input/output paths, model
- `src/docx_preserve.py` — extract/replace text units while keeping layout/images
- `src/markdown_to_docx.py` — simple Markdown → DOCX (not used in main flow)
- `input_documents/` — put source `.docx` here
- `output_documents/` — translated `.docx` here
- `.env` — API keys (`GOOGLE_API_KEY`, `GEMINI_API_KEY`)
- `requirements.txt` — dependencies

## Requirements
- Python 3.10+
- A Google AI Studio key (Gemini)  
  Get one: https://aistudio.google.com/app/apikey

## Setup
```bash
# 1) Create & activate venv (PowerShell)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2) Install deps
python -m pip install --upgrade pip
pip install -r requirements.txt

# 3) Add .env with your key
# In project root create .env with:
# GOOGLE_API_KEY=YOUR_KEY
# GEMINI_API_KEY=YOUR_KEY
```

## Configuration
Edit `config.yaml`:
```yaml
translation:
  target_language: "Hindi"     # e.g., "Hindi", "Spanish", "French"
llm:
  model_name: "gemini/gemini-1.5-flash"  # or gemini/gemini-1.5-pro
paths:
  input_file: "input_documents/academic_paper.docx"
  output_file: "output_documents/translated_paper.docx"
```

## Run
```bash
# Activate venv first
.\.venv\Scripts\Activate.ps1

# Execute
python run.py
```
- Output is saved to `output_documents/translated_paper.docx`.
- If that file is open/locked, a timestamped fallback name is used automatically.

## Change Target Language
- Set `translation.target_language` in `config.yaml` (e.g., `"Hindi"`).
- Re-run `python run.py`.

