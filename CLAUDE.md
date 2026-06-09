# CLAUDE.md

Implementer-side notes for this repo. README.md covers what the tool does for end users — this file covers what is non-obvious when **changing** the code.

## Stack

- .NET 8, single-file publish, `RuntimeIdentifier = win-x64` (set in `src/DocxTemplateEngine/DocxTemplateEngine.csproj`).
- `DocumentFormat.OpenXml` 3.2 — DOCX/OOXML manipulation.
- `Markdig` — Markdown parsing.
- `YamlDotNet` — YAML config support (JSON via `System.Text.Json`).
- `OpenMcdf` 3.1.3 — Compound File Binary Format (CFBF) for OLE Package wrapping.

Single-file publish constrains us: don't pull in `System.Drawing.Common` or other native-heavy deps. PNG/JPEG dimensions are read from binary headers in `ImageHandler.ReadImageDimensions` for this reason.

## Build / test

```bash
dotnet build                                                    # compile
dotnet test                                                     # all tests
dotnet test --filter "FullyQualifiedName~SomeTestName"          # one
dotnet publish src/DocxTemplateEngine -c Release                # single-file exe
```

## Code map

```
src/DocxTemplateEngine/
  Program.cs                      CLI arg parsing + entry point
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
    OlePackageWriter.cs           Builds the CFBF container for ProgID="Package" embeds
  Converters/                     Markdig -> OpenXML conversion helpers
tests/DocxTemplateEngine.Tests/
  TestDocxHelper.cs               Builds a minimal in-memory template docx for tests
  *Tests.cs                       xUnit + FluentAssertions
```

## Handler dispatch model

`TemplateProcessor._handlers` is a `Dictionary<PlaceholderType, IPlaceholderHandler>` (TemplateProcessor.cs:10). To add a new placeholder type:

1. Add the enum value in `Models/PlaceholderType.cs`.
2. Add a handler class in `Handlers/` implementing `IPlaceholderHandler.Replace(match, source, document, entry)`.
3. Register it in the `_handlers` dictionary in `TemplateProcessor`.
4. If the handler needs new fields, add them to `PlaceholderEntry` with both `[JsonPropertyName]` (JSON) and matching camelCase YAML name (the YAML deserializer uses `CamelCaseNamingConvention`).

Placeholders are processed in **reverse order** (TemplateProcessor.cs:69) so position-based mutations earlier in the body don't invalidate later matches.

## The split-run problem

Word frequently splits `{{Placeholder}}` across multiple `<w:r>` runs (e.g. autocorrect, formatting boundaries). `PlaceholderFinder.FindAll` walks paragraphs and reconstructs placeholder names across runs. **Always** use `PlaceholderFinder.RemovePlaceholderText(match)` from a handler rather than hand-rolling text removal — it returns the anchor run you should attach new content to.

## Path resolution + early validation

`TemplateConfig.Load`:

- Resolves every `entry.Source` **relative to the config file's directory**, not the process CWD.
- Calls `Validate()` which fails fast if any source file is missing or an image has an unsupported extension. By the time a handler runs, sources are guaranteed to exist on disk.

If you add a new placeholder type that has format constraints (e.g. only `.md` files), put the check in `TemplateConfig.Validate` so failures surface before any document mutation begins.

## File embedding (FileObjectHandler) — two paths

This is the trickiest part of the codebase. `FileObjectHandler.Replace` branches on the resolved `ProgID`:

| ProgID | Part type | Relationship | Payload |
|---|---|---|---|
| `Excel.Sheet.12`, `Word.Document.12`, `PowerPoint.Show.12`, ... | `EmbeddedPackagePart` | `package` | The file's raw bytes (the file *is* an OOXML package) |
| `Package` (everything else: pdf, txt, csv, yaml, json, xml, ...) | `EmbeddedObjectPart` | `oleObject` | A **CFBF container** built by `OlePackageWriter.Build`, with `Ole` / `CompObj` / `Ole10Native` streams wrapping the file |

If you feed Word an `EmbeddedPackagePart` containing raw non-OOXML bytes, the icon renders but **double-click silently does nothing** — Word can't extract a file from a non-package payload. The CFBF wrapper is what makes activation work for arbitrary file types.

The CFBF root storage **must have its CLSID set** to `{0003000C-0000-0000-C000-000000000046}` (the legacy Object Packager class). Word looks this CLSID up to find the OLE handler; without it, the icon renders but clicking it raises *"The server application, source file, or item cannot be found"*. See `OlePackageWriter.PackageClsid`.

The `Ole10Native` payload layout follows Apache POI's writer exactly (see the doc comment on `BuildOle10Native`). Two failure modes seen during this work:

- Forgetting the **length prefix** on the command/path field → click raises *"server application not found"* (Word can't even reach the data).
- Forgetting the **trailing `flags3` WORD** (`0x0000` after the file data) → click raises *"Embedded object is corrupted or has an empty file name"*.

Don't add a separate `tempPath` field — oletools tolerates one but Word's reader rejects the document. POI's two-string + length-prefixed-command + data + trailing-WORD layout is what Word actually expects.

Other invariants of the OLE structure (in `CreateOleObjectParagraph`):

- VML shape id is `_x0000_i####` (inline-OLE convention). The `_x0000_s####` form is for floating shapes and confuses Word here.
- The shape needs the `o:ole=""` attribute (set via `SetAttribute(new OpenXmlAttribute(...))`). Without it Word treats the shape as a plain image and never activates the OLE binding.
- The icon is a hardcoded base64 PNG embedded as `GenericFileIconPngBase64` — a real raster image, not a stub. A malformed image produces "The linked image cannot be displayed" in the rendered doc.
- `<w:object>` needs `DxaOriginal`/`DyaOriginal` (twentieths of a point). Some Word versions refuse to activate the OLE without these.

To add a new file extension, decide its ProgID:

- Native OOXML format (`.xlsx`, `.docx`, `.pptx`) → use the matching Office ProgID.
- Anything else → `Package`. Don't invent ProgIDs like `AcroExch.Document.DC` — they assume a specific app is installed for activation, whereas `Package` defers to the OS file association and works universally.

## OpenMcdf 3.x gotchas

These bit me when wiring up `OlePackageWriter`:

- `RootStorage.Commit()` throws `NotSupportedException("Cannot commit non-transacted storage")` unless the storage was created with `StorageModeFlags.Transacted`. The default mode writes through to the underlying stream as you go — no Commit needed. See `OlePackageWriter.Build`.
- CFBF system stream names start with the literal byte `0x01`. The actual on-disk name for the OLE 1.0 native data stream is the byte sequence `01 4F 6C 65 31 30 4E 61 74 69 76 65` (`0x01` followed by `Ole10Native`). OpenMcdf does **not** prepend this byte for you — pass it as part of the string. In `OlePackageWriter.cs` the constants `OleStreamName`, `CompObjStreamName`, `Ole10NativeStreamName` are written with a leading `` escape and already include the prefix.
- The stream returned by `OpenXmlPart.GetStream()` is **not seekable**, but `RootStorage.Open` requires a seekable stream. To read a CFBF embedded part in a test, `CopyTo` a `MemoryStream` first.
- The Read/Edit tools render `0x01` as invisible whitespace, so `"Ole"` displays identically to `"Ole"`. If an `Edit` whose `old_string` references one of these names fails to match, the in-file string almost certainly has the `0x01` prefix that your search string lacks. Use `xxd` to confirm, then either include the literal byte or rewrite the file with `Write`.

## Test conventions

- xUnit + FluentAssertions throughout.
- Each test class creates a per-test temp dir in its constructor (`Path.Combine(Path.GetTempPath(), $"DocxTests_{Guid.NewGuid():N}")`) and deletes it in `Dispose`. Drop new fixtures into the temp dir, not `tests/.../TestData/`.
- `TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "Foo")` builds a minimal docx containing `{{Foo}}`. Pass multiple names for multiple placeholders.
- For file-embedding tests, **assert the binary structure**, not just that an embedded part exists. The pre-fix bug passed `EmbeddedPackageParts.Should().NotBeEmpty()` while producing a broken document. Good assertions:
  - CFBF magic bytes `D0 CF 11 E0 A1 B1 1A E1` at offset 0 of the embedded object part.
  - Original file content present in the `Ole10Native` stream (round-trip check).
  - Icon image part `Length > 100` (catches stub-image regressions).

## Things that look like bugs but aren't

- The output docx is created by **copying the template file**, then mutating in place (TemplateProcessor.cs:30). This means the output preserves all template styling, headers, footers, custom XML — that's deliberate.
- Headers and footers are searched for placeholders in addition to the body (TemplateProcessor.cs:44–62). New handlers don't need extra wiring for this — the dispatch is body-agnostic.
- `MarkdownTableHandler` auto-detects whether the placeholder is **inside an existing `w:tbl`** and switches behavior (skip header row, append data rows, preserve template cell properties). Trigger: `FindAncestor<TableCell>(match.Paragraph)` returning non-null (MarkdownTableHandler.cs:17). If it's in a cell, `PopulateExistingTable` runs; otherwise `CreateNewTable`.
