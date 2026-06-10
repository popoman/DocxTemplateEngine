# DOCX Template Engine — Usage

A CLI tool that fills DOCX templates with content from markdown files, images, embedded file objects, and markdown tables. This document covers everything an end-user needs to drive `DocxTemplateEngine.exe`. For build / contribute instructions, see the project README on GitHub.

## Features

- **Markdown content** — Insert rich markdown (headings, bold, italic, lists, code blocks, links) into DOCX, preserving formatting.
- **Inline images** — Insert PNG/JPEG images at placeholder locations with configurable dimensions.
- **Embedded file objects** — Embed files (Excel, Word, PowerPoint, PDF, text, CSV, ZIP, ...) as OLE objects with clickable icons that activate via the OS file association.
- **Markdown tables** — Convert markdown pipe tables into native DOCX tables with styled headers.
- **Populate existing tables** — Place `{{Placeholder}}` inside a template table to fill it with markdown data rows, preserving your custom styling.
- **Split-run handling** — Correctly detects `{{Placeholders}}` even when Word splits them across multiple XML runs.
- **Dry-run mode** — Validate templates and configs without generating output.

## Running

The exe is self-contained except for the .NET 8 runtime. Place `DocxTemplateEngine.exe` somewhere on your PATH (or invoke it by full path) and run:

```bash
DocxTemplateEngine.exe -t template.docx -c config.json -o output.docx
```

### CLI Options

| Option | Short | Required | Description |
|--------|-------|----------|-------------|
| `--template` | `-t` | Yes | Path to the DOCX template file |
| `--config` | `-c` | Yes | Path to the JSON or YAML configuration file |
| `--output` | `-o` | Yes\* | Path for the generated output DOCX |
| `--verbose` | `-v` | No | Enable verbose logging |
| `--dry-run` | — | No | Validate config and template without generating output |
| `--help` | `-h` | No | Show help message |

\* Not required when using `--dry-run`.

## Configuration

Create a JSON or YAML configuration file that maps placeholder names to data sources. The format is auto-detected by file extension (`.json`, `.yaml`, or `.yml`).

### JSON format

```json
{
  "placeholders": {
    "Introduction": {
      "type": "markdown",
      "source": "content/intro.md"
    },
    "CompanyLogo": {
      "type": "image",
      "source": "assets/logo.png",
      "widthCm": 5.0,
      "heightCm": 3.0
    },
    "BudgetReport": {
      "type": "file",
      "source": "data/budget.xlsx",
      "displayName": "Budget Report Q1"
    },
    "MetricsTable": {
      "type": "markdownTable",
      "source": "data/metrics.md"
    }
  }
}
```

### YAML format

```yaml
placeholders:
  Introduction:
    type: Markdown
    source: content/intro.md
  CompanyLogo:
    type: Image
    source: assets/logo.png
    widthCm: 5.0
    heightCm: 3.0
  BudgetReport:
    type: File
    source: data/budget.xlsx
    displayName: Budget Report Q1
  MetricsTable:
    type: MarkdownTable
    source: data/metrics.md
```

### Placeholder Types

#### `markdown`

Reads a markdown file and converts it to DOCX-formatted content.

| Field | Required | Description |
|-------|----------|-------------|
| `type` | Yes | `"markdown"` |
| `source` | Yes | Path to the `.md` file |

**Supported markdown features:**
- Headings (H1–H6)
- Bold (`**text**`) and italic (`*text*`)
- Inline code (`` `code` ``) and fenced code blocks
- Ordered and unordered lists (including nested)
- Links (`[text](url)`)
- Block quotes
- Horizontal rules

#### `image`

Inserts an image inline at the placeholder location.

| Field | Required | Description |
|-------|----------|-------------|
| `type` | Yes | `"image"` |
| `source` | Yes | Path to a `.png` or `.jpeg` file |
| `widthCm` | No | Image width in centimeters |
| `heightCm` | No | Image height in centimeters |

If dimensions are not specified, the image's native dimensions are used (constrained to 16cm max width).

#### `file`

Embeds a file as an OLE object with an icon representation. Office formats (`.docx`, `.doc`, `.xlsx`, `.xls`, `.pptx`, `.ppt`) activate via the corresponding Office app; everything else activates via the OS file association (whatever the recipient has set as the default app for that extension).

| Field | Required | Description |
|-------|----------|-------------|
| `type` | Yes | `"file"` |
| `source` | Yes | Path to the file to embed |
| `displayName` | No | Display name for the icon (defaults to filename) |

Any file type can be embedded. Common ones tested: `.docx`, `.xlsx`, `.pptx`, `.pdf`, `.txt`, `.csv`, `.json`, `.yaml`, `.zip`.

#### `markdownTable`

Reads a markdown pipe table and converts it to a native DOCX table.

| Field | Required | Description |
|-------|----------|-------------|
| `type` | Yes | `"markdownTable"` |
| `source` | Yes | Path to a `.md` file containing a pipe table |

The header row is automatically styled with bold text and background shading.

**Auto-detect: Populating an existing table**

If the `{{Placeholder}}` is placed inside a cell of an **existing** table in the template, the handler automatically switches to "populate" mode:
- The markdown header row is **skipped** (the template table already has its own styled header).
- Only data rows from the markdown are appended to the existing table.
- Cell widths and properties from the template row are preserved.
- The row containing the placeholder is removed after data insertion.

This lets you design your table styling (borders, colors, column widths, header formatting) in Word/LibreOffice and have the data filled in from markdown.

**Example template table:**

| Name | Score |
|------|-------|
| {{DataTable}} | |

After processing with a markdown file containing `| Name | Score |\n|---|---|\n| Alice | 95 |\n| Bob | 87 |`, the placeholder row is replaced with the two data rows, keeping the styled header intact.

### Path Resolution

All `source` paths are resolved relative to the **config file's** directory. Absolute paths are also supported.

## Template Format

Create your DOCX template in any word processor (Word, LibreOffice, etc.) and use `{{PlaceholderName}}` markers where content should be inserted.

**Example template content:**

```
Project Report

{{Introduction}}

Company Logo: {{CompanyLogo}}

Attached Budget: {{BudgetReport}}

Performance Metrics:
{{MetricsTable}}
```

> **Tip:** The engine handles the "split-run" problem where Word editors may internally split `{{PlaceholderName}}` across multiple XML runs. No special template preparation is needed.
