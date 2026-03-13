# AI Paper Review and Response Generator

An automated system that uses Large Language Models (LLMs) to generate peer-review comments for academic papers and draft the corresponding author responses — then exports everything as formatted Word documents.

---

## 主要功能 / Main Features

| Feature | Description |
|---|---|
| **PDF Parsing** | Extracts the full text content from an academic PDF file |
| **Discipline Detection** | Automatically identifies the academic field of the paper (e.g., Computer Science, Biology, Economics) |
| **Review Generation** | Produces rigorous, problem-focused peer-review comments anchored to specific sections of the paper |
| **Response Generation** | Drafts professional author responses to each reviewer comment |
| **DOCX Export** | Saves review comments and author responses as two separate, formatted Word documents |
| **Bilingual Support** | Generates output in any language (English, Chinese, etc.) via the `--language` flag |
| **Custom Formatting** | Configurable document style (fonts, margins, spacing) through JSON format presets |
| **Dual-Model Support** | Allows using different LLM models and temperature settings for the reviewer and responder roles |

---

## 实现方法 / Implementation

### Technology Stack

- **Python 3** — CLI application with no web framework dependency
- **[PyMuPDF (fitz)](https://pymupdf.readthedocs.io/)** — PDF text extraction
- **[python-docx](https://python-docx.readthedocs.io/)** — Word document generation
- **Standard library `urllib`** — HTTP client for LLM API calls (no external HTTP library required)

### Architecture

The pipeline runs four sequential steps:

```
[1/4] PDF Parser
      └─ Extracts raw text from the input PDF file

[2/4] Discipline Detector
      └─ Sends a sample of the paper text to the LLM and asks it to name the academic field

[3/4] Review Generator  ← primary LLM (--model)
      └─ Prompts the LLM as a rigorous peer reviewer
         Returns a JSON array: [{reference location, reference text, issue, detail}, ...]

[4/4] Review Responder  ← responder LLM (--responder-model, defaults to --model)
      └─ Prompts the LLM as the paper's author
         Returns a JSON array: [{problem, responde}, ...]
```

### Module Overview

| File | Role |
|---|---|
| `start.py` | Entry point — argument parsing and pipeline orchestration |
| `pdf_parser.py` | PDF text extraction using PyMuPDF |
| `llm_client.py` | OpenAI-compatible API client |
| `discipline_detector.py` | LLM-powered academic discipline detection |
| `review_generator.py` | LLM-powered peer-review question generation |
| `review_responder.py` | LLM-powered author response generation |
| `docx_exporter.py` | Word document export with configurable formatting |
| `format_zh_academic.json` | Chinese academic document style preset |
| `format_en_academic.json` | English academic document style preset |

### LLM Integration

The system works with **any OpenAI-compatible API endpoint** — including OpenAI, Azure OpenAI, and locally hosted models. All prompts are engineered to return **structured JSON arrays**, with built-in fallback parsing to handle code-fenced or otherwise wrapped LLM output.

---

## 安装 / Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/aaron-bai/AI-paper-review-and-response-generator.git
   cd AI-paper-review-and-response-generator
   ```

2. **Install dependencies**

   ```bash
   pip install pymupdf python-docx
   ```

   > Python 3.10 or later is recommended.

---

## 使用方法 / Usage

Run `start.py` from the terminal and pass the required arguments:

```bash
python start.py \
    --pdf         path/to/paper.pdf \
    --api-url     https://api.openai.com/v1 \
    --api-token   sk-... \
    --model       gpt-4o \
    --output-dir  ./output
```

### Required Arguments

| Argument | Description |
|---|---|
| `--pdf` | Path to the academic paper PDF file |
| `--api-url` | Base URL of the OpenAI-compatible LLM API |
| `--api-token` | API key / token for the LLM service |
| `--model` | Primary LLM model name (used for discipline detection and review generation) |
| `--output-dir` | Directory where the output Word documents will be saved |

### Optional Arguments

| Argument | Default | Description |
|---|---|---|
| `--temperature` | `0.7` | Sampling temperature for the primary LLM (`0.0` = deterministic, `1.0` = most creative) |
| `--responder-model` | same as `--model` | LLM model used for generating author responses |
| `--responder-temperature` | same as `--temperature` | Sampling temperature for the responder LLM |
| `--language` | `English` | Language for all generated text (e.g., `English`, `Chinese`) |
| `--review-aspects` | `methodology` | Aspects to focus on during review (newline-separated list, e.g., `methodology`, `novelty`, `clarity`) |
| `--format` | `zh` | DOCX style preset: `zh` (Chinese academic), `en` (English academic), or a path to a custom JSON file |
| `--output-path` | *(none)* | Optional path for a combined plain-text output file |

### Output Files

After a successful run the following files are written to `--output-dir`:

| File | Contents |
|---|---|
| `评审意见_<timestamp>.docx` | Formatted peer-review comments |
| `回复_<timestamp>.docx` | Formatted author responses |

If `--output-path` is provided, a plain-text version of both documents is also written to that path.

---

## 示例 / Examples

### English review with a single model

```bash
python start.py \
    --pdf       my_paper.pdf \
    --api-url   https://api.openai.com/v1 \
    --api-token sk-xxxxxxxxxxxxxxxx \
    --model     gpt-4o \
    --language  English \
    --format    en \
    --output-dir ./reviews
```

### Chinese review with separate reviewer and responder models

```bash
python start.py \
    --pdf                  paper.pdf \
    --api-url              https://api.openai.com/v1 \
    --api-token            sk-xxxxxxxxxxxxxxxx \
    --model                gpt-4o \
    --temperature          0.8 \
    --responder-model      gpt-4o-mini \
    --responder-temperature 0.3 \
    --language             Chinese \
    --format               zh \
    --output-dir           ./output \
    --output-path          combined_output.txt
```

### Focus on specific review aspects

```bash
python start.py \
    --pdf            paper.pdf \
    --api-url        https://api.openai.com/v1 \
    --api-token      sk-xxxxxxxxxxxxxxxx \
    --model          gpt-4o \
    --review-aspects $'methodology\nnovelty\nstatistical analysis' \
    --output-dir     ./output
```

---

## 自定义格式 / Custom Formatting

The document style is controlled by a JSON configuration file. Two presets are included:

- `format_zh_academic.json` — Chinese academic style (宋体 / Times New Roman, standard Chinese margins)
- `format_en_academic.json` — English academic style

To apply a custom style, create your own JSON file following the same schema and pass its path to `--format`:

```bash
python start.py ... --format /path/to/my_format.json
```

Key configurable fields:

```jsonc
{
  "document": {
    "default_font": { "latin": "Times New Roman", "east_asia": "宋体", "size": 12 },
    "margins_cm": { "top": 2.54, "bottom": 2.54, "left": 3.18, "right": 3.18 }
  },
  "paragraph": { "line_spacing": 1.5, "space_after_pt": 6, "first_line_indent_chars": 2 },
  "title": { "font_size": 16, "bold": true, "alignment": "center" },
  "section_heading": { "font_size": 14, "bold": true },
  "review": {
    "labels": {
      "reference_location": "Reference Location",
      "reference_text": "Reference Text",
      "issue": "Issue",
      "detail": "Detail"
    }
  },
  "response": {
    "labels": { "problem": "Problem", "responde": "Response" }
  }
}
```

---

## License

[MIT License](LICENSE) — Copyright © 2026 Bai Yufan