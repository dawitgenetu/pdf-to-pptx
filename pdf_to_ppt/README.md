# PDF → PowerPoint Generator

Automatically converts any PDF into a professionally designed `.pptx` presentation.

## Features
- Multi-page PDF parsing (PyMuPDF)
- NLP-based section detection, summarisation & keyword extraction (spaCy)
- Clean blue/white theme with accent colours
- Keyword highlighting in slides (bold + orange)
- Auto-splits long sections across multiple slides
- Streamlit web UI **and** CLI

---

## Setup

```bash
cd pdf_to_ppt
pip install -r requirements.txt
python -m spacy download en_core_web_sm
```

---

## Usage

### Streamlit UI
```bash
streamlit run app.py
```
Open http://localhost:8501, upload your PDF, click **Generate Presentation**, download the `.pptx`.

### CLI
```bash
python cli.py "path/to/file.pdf"
# or specify output path:
python cli.py "path/to/file.pdf" "output.pptx"
```

---

## Project Structure

```
pdf_to_ppt/
├── pdf_parser.py      # PDF text extraction & cleaning
├── nlp_processor.py   # Section detection, summarisation, keywords
├── ppt_generator.py   # PowerPoint slide building
├── app.py             # Streamlit web UI
├── cli.py             # Command-line interface
└── requirements.txt
```

## Slide Types Generated
| Slide | Description |
|-------|-------------|
| Title slide | PDF title + subtitle |
| Section header | One per detected section |
| Content slides | 5–6 bullet points, keywords highlighted |
| Summary slide | Table of contents |
