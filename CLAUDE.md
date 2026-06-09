# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Implementer-side notes. README.md covers what the tool does for end users — this file covers what is non-obvious when **changing** the code.

## Stack

- .NET 8, single-file publish, `RuntimeIdentifier = win-x64` (set in `src/DocxTemplateEngine/DocxTemplateEngine.csproj`).
- `DocumentFormat.OpenXml` 3.2 — DOCX/OOXML manipulation.
- `Markdig` — Markdown parsing.
- `YamlDotNet` — YAML config support (JSON via `System.Text.Json`).
- `OpenMcdf` 3.1.3 — declared but **currently unused** (see "OLE embedding limitation" below).

Single-file publish constrains us: don't pull in `System.Drawing.Common` or other native-heavy deps. PNG/JPEG dimensions are read from binary headers in `ImageHandler.ReadImageDimensions` for this reason.

## Build / test / run

```bash
dotnet build                                                          # compile
dotnet test                                                           # all tests
dotnet test --filter "FullyQualifiedName~SomeTestName"                # one
dotnet publish src/DocxTemplateEngine -c Release                      # single-file exe

# Run the CLI directly against a config:
dotnet run --project src/DocxTemplateEngine -- \
    -t template.docx -c config.yaml -o out.docx [-v] [--dry-run]
```

End-to-end smoke test (`invoke/build-document.ps1`): finds the freshest local Debug build of `DocxTemplateEngine.exe`, closes any Word instance holding `res.docx`, deletes the stale output, regenerates from `invoke/template.docx` + `invoke/config.yaml`, and opens the result. Exercises every handler type (markdown, image, markdownTable, file with .pdf/.yaml/.docx). Use this whenever you touch a handler — the unit tests check structure, this checks "does Word actually like the file."

## Code map

```
src/DocxTemplateEngine/
  Program.cs                      CLI arg parsing + entry point
  Resources/app.ico                Application icon (referenced by .csproj)
  Engine/
    TemplateProcessor.cs          Orchestrates: opens docx, finds placeholders, dispatches to handlers, saves
    PlaceholderFinder.cs          Finds {{Name}} across split Word runs; provides RemovePlaceholderText
  Models/
    PlaceholderType.cs            Enum: Markdown | Image | File | MarkdownTable
    PlaceholderEntry.cs           One config entry (type, source, displayName, widthCm, heightCm)
    TemplateConfig.cs             Loads + validates JSON/YAML config
  Handlers/
    IPlaceholderHandler.cs        Single-method interface (Replace)
    MarkdownHandler.cs            Markdown -> OpenXML elements
    MarkdownTableHandler.cs       Pipe table -> w:tbl, with auto-detect of "populate existing table" mode
    ImageHandler.cs               Inline image insertion
    FileObjectHandler.cs          OLE-embedded file with clickable icon
  Converters/                     Markdig -> OpenXML conversion helpers
tests/DocxTemplateEngine.Tests/
  TestDocxHelper.cs               Builds a minimal in-memory template docx for tests
  *Tests.cs                       xUnit + FluentAssertions
invoke/                           Smoke-test harness (PS1 + sample template/config/assets)
```

## Handler dispatch model

`TemplateProcessor._handlers` is a `Dictionary<PlaceholderType, IPlaceholderHandler>` (TemplateProcessor.cs:10). To add a new placeholder type:

1. Add the enum value in `Models/PlaceholderType.cs`.
2. Add a handler class in `Handlers/` implementing `IPlaceholderHandler.Replace(match, source, document, entry)`.
3. Register it in the `_handlers` dictionary in `TemplateProcessor`.
4. If the handler needs new fields, add them to `PlaceholderEntry` with both `[JsonPropertyName]` (JSON) and matching camelCase YAML name (the YAML deserializer uses `CamelCaseNamingConvention`).

Placeholders are processed in **reverse order** (TemplateProcessor.cs:69) so position-based mutations earlier in the body don't invalidate later matches. The dispatch is body-agnostic — headers and footers are scanned (TemplateProcessor.cs:44–62) and feed the same handler list, so new handlers don't need extra wiring for those.

## The split-run problem

Word frequently splits `{{Placeholder}}` across multiple `<w:r>` runs (e.g. autocorrect, formatting boundaries). `PlaceholderFinder.FindAll` walks paragraphs and reconstructs placeholder names across runs. **Always** use `PlaceholderFinder.RemovePlaceholderText(match)` from a handler rather than hand-rolling text removal — it returns the anchor run you should attach new content to. `TestDocxHelper.CreateTemplateWithSplitRuns` builds a template with the placeholder pre-split into three runs to regression-test this.

## Path resolution + early validation

`TemplateConfig.Load`:

- Resolves every `entry.Source` **relative to the config file's directory**, not the process CWD.
- Calls `Validate()` which fails fast if any source file is missing or an image has an unsupported extension. By the time a handler runs, sources are guaranteed to exist on disk.

If you add a new placeholder type that has format constraints (e.g. only `.md` files), put the check in `TemplateConfig.Validate` so failures surface before any document mutation begins.

## File embedding (FileObjectHandler) — current state and known limitation

`FileObjectHandler.Replace` does the same thing for **every** file extension:

1. `mainPart.AddEmbeddedPackagePart(contentType)` — adds an `/word/embeddings/.../*` part with a `package` relationship.
2. `embeddedPart.FeedData(fileStream)` — writes the file's **raw bytes** into that part.
3. Adds an EMF icon image (built by `GenerateMinimalIconEmf` — a literal hand-rolled minimal EMF, not a real raster) and a VML shape + `OleObject` element pointing at both the icon and the embedded part.

`ProgIdMap` (FileObjectHandler.cs:17) selects a ProgID per extension: `Excel.Sheet.12` for `.xlsx`/`.csv`, `Word.Document.12` for `.docx`, `AcroExch.Document.DC` for `.pdf`, `Package` for `.txt`/`.zip`, etc.

**Known limitation — non-OOXML embeds.** Word expects non-OOXML embeds (PDF, TXT, CSV, JSON, etc.) to be wrapped in a CFBF (Compound File Binary Format) container with `Ole`/`CompObj`/`Ole10Native` streams and the root storage CLSID set to `{0003000C-0000-0000-C000-000000000046}` (legacy Object Packager). The current code skips that and just stores raw bytes. The icon renders, but double-click activation is unreliable: behavior ranges from "opens fine" (when the host app handles the ProgID directly) to "silent no-op" or *"server application, source file, or item cannot be found"*. The `OpenMcdf` package was added in anticipation of writing a CFBF wrapper but the wrapper is not yet implemented.

The current integration test (`TemplateProcessorIntegrationTests.cs:186`) only checks `EmbeddedPackageParts.Should().NotBeEmpty()` — that asserts a part exists, not that the embed actually activates in Word. If you implement CFBF wrapping, replace this with structural assertions: CFBF magic bytes `D0 CF 11 E0 A1 B1 1A E1` at offset 0, original file content round-tripped through the `Ole10Native` stream, etc., and load the output in real Word as part of the smoke test (`invoke/build-document.ps1`).

When picking a ProgID for a new extension: native OOXML formats use the matching Office ProgID. For everything else, prefer `Package` (defers to OS file association, works without any specific app installed) over app-specific ProgIDs like `AcroExch.Document.DC` which silently fail when the assumed app isn't on the target machine.

## Test conventions

- xUnit + FluentAssertions throughout.
- Each test class with file I/O implements `IDisposable`, creates a per-test temp dir in its constructor (`Path.Combine(Path.GetTempPath(), $"DocxTests_{Guid.NewGuid():N}")`), and deletes it in `Dispose`. Drop new fixtures into the temp dir, not `tests/.../TestData/`.
- `TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "Foo")` builds a minimal docx containing `{{Foo}}`. Pass multiple names for multiple placeholders. Variants exist for split-run and mixed-content templates.
- Structural assertions beat existence assertions — see the `EmbeddedPackageParts.Should().NotBeEmpty()` example above for what *not* to settle for. For embeds, prefer asserting bytes/content; for tables, assert row counts and cell text; for images, assert the `Drawing` element + image part length.

## Things that look like bugs but aren't

- The output docx is created by **copying the template file**, then mutating in place (TemplateProcessor.cs:30). This means the output preserves all template styling, headers, footers, custom XML — that's deliberate.
- Headers and footers are searched for placeholders in addition to the body. New handlers don't need extra wiring for this — the dispatch is body-agnostic.
- `MarkdownTableHandler` auto-detects whether the placeholder is **inside an existing `w:tbl`** and switches behavior (skip header row, append data rows, preserve template cell properties). Trigger: `FindAncestor<TableCell>(match.Paragraph)` returning non-null (MarkdownTableHandler.cs:17). If it's in a cell, `PopulateExistingTable` runs; otherwise `CreateNewTable`.
