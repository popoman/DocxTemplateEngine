# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Implementer-side notes. README.md covers what the tool does for end users — this file covers what is non-obvious when **changing** the code.

## Stack

- .NET 8, single-file publish, `RuntimeIdentifier = win-x64` (set in `src/DocxTemplateEngine/DocxTemplateEngine.csproj`).
- `DocumentFormat.OpenXml` 3.2 — DOCX/OOXML manipulation.
- `Markdig` — Markdown parsing.
- `YamlDotNet` — YAML config support (JSON via `System.Text.Json`).
- `OpenMcdf` 3.1.3 — builds the CFBF compound-file container used to wrap non-Office embeds (see "File embedding" below).
- `System.Text.Encoding.CodePages` — required to use windows-1252 (CP 1252) under single-file publish; `OlePackageBuilder` registers the provider in a static initializer.

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
    FileObjectHandler.cs          OLE-embedded file with clickable icon (Office vs generic CFBF path)
    OlePackageBuilder.cs          CFBF Object Packager wrapper for non-Office embeds
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

## File embedding (FileObjectHandler)

`FileObjectHandler.Replace` branches on extension via `OfficeExtensions` (`.docx`, `.doc`, `.xlsx`, `.xls`, `.pptx`, `.ppt`):

- **Office path:** `mainPart.AddEmbeddedPackagePart(contentType)` + `FeedData(fileStream)` writes raw OOXML bytes; ProgID comes from `OfficeProgIdMap` (e.g. `Word.Document.12`). Word activates these natively.
- **Generic path (everything else):** `mainPart.AddEmbeddedObjectPart("application/vnd.openxmlformats-officedocument.oleObject")` for an `oleObject` relationship; bytes come from `OlePackageBuilder.Build(source, fileName)` which produces a CFBF compound file. ProgID is hard-coded to `Package` so activation defers to the OS file association (works without Acrobat/Notepad++/etc. specifically installed).

Both paths share `CreateOleObjectParagraph` for the VML shape + `OleObject` element + icon image part.

### `OlePackageBuilder` — the CFBF blob

Reproduces, byte-for-byte, what Word writes when you "Insert > Object > Create from file > Display as icon" on a non-Office file. Root storage CLSID `{0003000C-0000-0000-C000-000000000046}` (Object Packager) plus three streams:

| Stream | Purpose | Notes |
|---|---|---|
| `\x01CompObj` | Class identification | 76 bytes, fixed except for CLSID. AnsiUserType `"OLE Package"`, AnsiClipboardFormat absent (marker = 0), Reserved1 string `"Package"`, Unicode marker `0x71B239F4`, all unicode tail strings empty. |
| `\x03ObjInfo` | Word-specific object state | Static 6 bytes: `40 00 03 00 01 00`. Note the `\x03` prefix (not `\x01`) per [MS-DOC] ObjectPool conventions. |
| `\x01Ole10Native` | The wrapped file | uint32 streamSize, uint16 `0x0002`, AnsiZ label, AnsiZ originalPath, uint32 `0x00030000`, uint32 tempPathLen, AnsiZ tempPath, uint32 nativeDataLen, raw file bytes. |

Stream-name prefixes (`\x01`, `\x03`) are control bytes embedded directly in the C# string literals. The Read tool and most editors hide them, so search for the const declarations by name (`CompObjStreamName` etc.) rather than by visible text. OpenMcdf 3.1.3 stores stream names verbatim — it does **not** auto-prepend the prefix.

`OlePackageBuilder` writes via windows-1252 (`Encoding.GetEncoding(1252)`). The static initializer calls `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)` because CP 1252 is not built into the trimmed BCL used by single-file publish.

### Adding a new extension

1. Office-native (matched by Office app's OOXML format) → add to `OfficeExtensions` + `OfficeProgIdMap` + `OfficeContentTypeMap`.
2. Anything else → no code change needed. The generic path handles arbitrary binaries; ProgID `Package` + the OS file association will do the right thing on the target machine.

Do not bring back app-specific ProgIDs (`AcroExch.Document.DC`, `Excel.Sheet.12` for `.csv`, etc.) — they silently fail when the assumed host app isn't installed or registered.

### Verifying changes

`Process_FileObjectPlaceholder_EmbedsFile` and `Process_FileObjectPlaceholder_OfficeFile_StoresRawPackage` (both in `TemplateProcessorIntegrationTests.cs`) cover the structural side: CFBF magic, root CLSID, stream presence, label and payload round-trip for the generic path, and ZIP magic preservation for the Office path. They will not catch behavioral regressions in real Word — run `invoke/build-document.ps1` and double-click each icon to confirm activation.

## Test conventions

- xUnit + FluentAssertions throughout.
- Each test class with file I/O implements `IDisposable`, creates a per-test temp dir in its constructor (`Path.Combine(Path.GetTempPath(), $"DocxTests_{Guid.NewGuid():N}")`), and deletes it in `Dispose`. Drop new fixtures into the temp dir, not `tests/.../TestData/`.
- `TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "Foo")` builds a minimal docx containing `{{Foo}}`. Pass multiple names for multiple placeholders. Variants exist for split-run and mixed-content templates.
- Structural assertions beat existence assertions. `EmbeddedPackageParts.Should().NotBeEmpty()` proves a part exists, not that it activates — the OLE-embed test now asserts CFBF magic, root CLSID, stream names, and round-trips the payload through `Ole10Native`. Apply the same standard elsewhere: for tables assert row counts and cell text; for images assert the `Drawing` element + image part length.

## Things that look like bugs but aren't

- The output docx is created by **copying the template file**, then mutating in place (TemplateProcessor.cs:30). This means the output preserves all template styling, headers, footers, custom XML — that's deliberate.
- Headers and footers are searched for placeholders in addition to the body. New handlers don't need extra wiring for this — the dispatch is body-agnostic.
- `MarkdownTableHandler` auto-detects whether the placeholder is **inside an existing `w:tbl`** and switches behavior (skip header row, append data rows, preserve template cell properties). Trigger: `FindAncestor<TableCell>(match.Paragraph)` returning non-null (MarkdownTableHandler.cs:17). If it's in a cell, `PopulateExistingTable` runs; otherwise `CreateNewTable`.
