# DOCX Template Engine

A .NET 8 CLI tool that fills DOCX templates with content from markdown files, images, embedded file objects, and markdown tables.

> **End-user documentation** lives in [`src/DocxTemplateEngine/DocxTemplateEngine.usage.md`](src/DocxTemplateEngine/DocxTemplateEngine.usage.md). It covers CLI flags, the JSON/YAML config schema, every placeholder type, and the template format. The same file is copied next to `DocxTemplateEngine.exe` on every build and publish, so users always have it on hand. This README is for developers cloning the repo.

## Features

- Markdown → DOCX with styles preserved
- Inline PNG/JPEG images with optional sizing
- Embedded file objects (OLE) for any file type, with clickable icons that activate via the host's default app
- Markdown pipe tables → native DOCX tables, including a "populate existing table" mode that preserves your custom styling
- Robust to Word's run-splitting of `{{Placeholders}}`
- `--dry-run` for config/template validation

## Build

```bash
dotnet build
```

## Publish (single-file exe)

```bash
dotnet publish src/DocxTemplateEngine -c Release
```

Produces a single `DocxTemplateEngine.exe` (~8 MB) in `src/DocxTemplateEngine/bin/Release/net8.0/win-x64/publish/`, with `DocxTemplateEngine.usage.md` copied alongside it. Requires the .NET 8 runtime on the target machine.

The csproj pins `RuntimeIdentifier` to `win-x64`. To target another platform, override on the command line:

```bash
dotnet publish src/DocxTemplateEngine -c Release -r linux-x64
dotnet publish src/DocxTemplateEngine -c Release -r osx-x64
```

## Run from source

```bash
dotnet run --project src/DocxTemplateEngine -- \
  --template template.docx \
  --config config.json \
  --output output.docx
```

## Tests

```bash
dotnet test
```

There is also a smoke harness at `invoke/build-document.ps1` that exercises every handler type against `invoke/template.docx` + `invoke/config.yaml` and opens the result in Word — use it whenever you touch a handler.

## Project structure

```
DocxTemplateEngine/
├── src/DocxTemplateEngine/
│   ├── Program.cs                       # CLI entry point
│   ├── DocxTemplateEngine.usage.md      # end-user doc, copied to publish output
│   ├── Models/                          # Config models and enums
│   ├── Engine/                          # Template processor and placeholder finder
│   ├── Handlers/                        # Type-specific placeholder handlers
│   │   ├── FileObjectHandler.cs         # OLE-embedded files (Office vs generic CFBF path)
│   │   ├── OlePackageBuilder.cs         # CFBF Object Packager wrapper for non-Office embeds
│   │   ├── ImageHandler.cs
│   │   ├── MarkdownHandler.cs
│   │   └── MarkdownTableHandler.cs
│   ├── Converters/                      # Markdig → OpenXml conversion helpers
│   └── Resources/                       # App icon, bundled file-object icon
└── tests/DocxTemplateEngine.Tests/
    ├── *Tests.cs                        # xUnit + FluentAssertions
    └── TestDocxHelper.cs                # Test DOCX file generator
```

For implementer-side notes (handler dispatch model, the split-run problem, the OLE/CFBF byte layout), see `CLAUDE.md`.

## Dependencies

- [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) — DOCX manipulation
- [Markdig](https://github.com/xoofx/markdig) — Markdown parsing
- [YamlDotNet](https://github.com/aaubry/YamlDotNet) — YAML configuration support
- [OpenMcdf](https://github.com/ironfede/openmcdf) — CFBF compound-file container for non-Office OLE embeds
- `System.Text.Encoding.CodePages` — windows-1252 support under single-file publish

## License

MIT
