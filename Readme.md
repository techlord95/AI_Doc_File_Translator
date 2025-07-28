# Advanced Doc File Translator

- A modern, full-featured web application for translating DOCX and PDF documents into multiple languages, preserving formatting, lists, and structure. Powered by advanced LLM APIs (Google Gemini, Groq), with support for document summarization and PDF-to-Word conversion.

- It started as an idea for my internship  and now it is a full fledged project, I worked on it for months but the upper management was happy with their employees doing manual translations, this tool aims to reduce the manual workload and save countless of hours, give it a star ⭐ if it helps you.

---

## Features

- **Translate DOCX and PDF files** to a wide range of languages. -Hindi, Arabic, German, Japanese, Korean
- **Preserves formatting**: Lists, bullet points, numbering, and text structure, image positioning are retained.
- **Multiple translation engines**: Choose between Google Gemini and Groq LLMs.
- **Document summarization**: Generate a concise summary of your document using LLMs.
- **PDF to DOCX conversion**: Uses `pdf2docx` or `Aspose.PDF` for high-fidelity conversion.*Both work differently choose according to your needs*
- **Modern web UI**: Drag-and-drop upload, progress bar, logs, and download links.
- **Customizable translation tone**: Professional, easy to understand, scientific, etc. Professional works best for most kinds of docuemnts
- **Cache and API usage stats**: See API call counts and cache hits.
- **Multi model selection**:Groq and Gemini support several models , each has their own strengths, select the model according your needs.
- **Download translated documents and summaries**.

---

## Demo

<!-- ![screenshot or gif here, if available] -->
-**Main Page**
<img width="1897" height="978" alt="image" src="https://github.com/user-attachments/assets/ff06107f-df85-4755-94bd-dcd58895a123" />



---

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/techlord95/AI_Doc_File_Translator.git
cd AI_Doc_File_Translator
```

### 2. Install dependencies

It's recommended to use a virtual environment:

```bash
python -m venv venv
# On Unix/macOS:
source venv/bin/activate
# On Windows:
venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Set up API Keys

- **Google Gemini**: Set the `GEMINI_API_KEY` environment variable.
- **Groq**: Set the `GROQ_API_KEY` environment variable.



Or edit the top of `advanced_docx_translator.py`

### 4. Run the web server

```bash
python web_server.py
```

Visit [http://localhost:5000](http://localhost:5000) in your browser.

---

## Usage


### Web Interface

- Drag and drop a `.docx` or `.pdf` file.
- Select target language, translation engine, tone, and PDF conversion engine.
- Optionally, generate a summary or translate only the first page.
- Click "Translate Document".
- Download the translated file and/or summary when ready.

### Command Line

You can also use the advanced translator directly:

```bash
python advanced_docx_translator.py input.docx --output_file output_translated.docx --target_language fr --engine gemini
```

---



---

## Requirements

See `requirements.txt`:

```
flask
werkzeug
python-docx
groq
deep-translator
aspose-pdf
pdf2docx
googletrans
```

You may need to install system dependencies for `aspose-pdf` and `pdf2docx` (see their docs).

---

## Environment Variables

- `GEMINI_API_KEY` – Google Gemini API key (for LLM translation)
- `GROQ_API_KEY` – Groq API key (for LLM translation and summarization)

---

## Extending

- Add more translation engines by implementing new functions in `advanced_docx_translator.py` , aspose and pdf2docx sometimes dont work well.
- Add more tones, languages, or PDF conversion engines via the UI and backend.
- Improve error handling, logging, or add authentication for production use.

---



## Acknowledgements

- Cat and Dog videos on the internet (they do make my day).

- God (they might have made this project possible, cause in initial days this project was a lost cause).
