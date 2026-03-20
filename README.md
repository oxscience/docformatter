# DocFormatter

**Drag & drop script formatter for .docx files.**

A web-based tool that reformats Word documents with consistent styling — character colors, scene headings, stage directions, and more. Built for the "Hedda & Pico" audio drama script format.

## Features

- **Drag & Drop**: Upload .docx files via drag & drop or file picker
- **Auto-Formatting**: Applies consistent styling rules:
  - Title in 48pt bold uppercase
  - Scene headings in 13pt bold underlined uppercase
  - Character names in bold uppercase with role-specific colors
  - Stage directions in 9pt italic
  - Narrator text in italic, indented
  - Page numbers right-aligned
- **Character Colors**: Each role gets a unique color (Hedda = orange, Wendt = green, etc.)
- **Case Character**: Mark the episode's case character to highlight them in purple
- **Privacy**: Files are processed in memory and immediately deleted — nothing stored on server
- **Deploy Anywhere**: Railway-ready with Gunicorn

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Backend | Python, Flask 3.1, python-docx |
| Frontend | Vanilla HTML/CSS/JS (inline) |
| Production | Gunicorn |
| Deploy | Railway, Docker, or any Python host |

## Quick Start

```bash
git clone https://github.com/oxscience/docformatter.git
cd docformatter
python -m venv venv && source venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open `http://localhost:5009`.

## Deploy to Railway

```bash
# Railway auto-detects the Procfile/Gunicorn setup
railway up
```

## License

MIT
