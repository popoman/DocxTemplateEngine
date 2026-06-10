using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Engine;
using DocxTemplateEngine.Models;
using FluentAssertions;
using OpenMcdf;

namespace DocxTemplateEngine.Tests;

public class TemplateProcessorIntegrationTests : IDisposable
{
    private readonly string _tempDir;

    public TemplateProcessorIntegrationTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"DocxTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public void Process_MarkdownPlaceholder_ReplacesContent()
    {
        var templatePath = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "Content");
        var mdPath = Path.Combine(_tempDir, "content.md");
        File.WriteAllText(mdPath, "# Hello World\n\nThis is a **test** paragraph.");

        var configJson = $$"""
        {
          "placeholders": {
            "Content": {
              "type": "markdown",
              "source": "content.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var outputPath = Path.Combine(_tempDir, "output.docx");
        var config = TemplateConfig.Load(configPath);
        var processor = new TemplateProcessor(verbose: false);

        processor.Process(templatePath, config, outputPath);

        File.Exists(outputPath).Should().BeTrue();

        // Open and verify the placeholder was replaced
        using var doc = WordprocessingDocument.Open(outputPath, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        var text = body.InnerText;

        text.Should().Contain("Hello World");
        text.Should().Contain("test");
        text.Should().NotContain("{{Content}}");
    }

    [Fact]
    public void Process_TablePlaceholder_CreatesTable()
    {
        var templatePath = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "DataTable");
        var mdPath = Path.Combine(_tempDir, "table.md");
        File.WriteAllText(mdPath, "| Name | Score |\n|------|-------|\n| Alice | 95 |\n| Bob | 87 |");

        var configJson = $$"""
        {
          "placeholders": {
            "DataTable": {
              "type": "markdownTable",
              "source": "table.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var outputPath = Path.Combine(_tempDir, "output_table.docx");
        var config = TemplateConfig.Load(configPath);
        var processor = new TemplateProcessor(verbose: false);

        processor.Process(templatePath, config, outputPath);

        File.Exists(outputPath).Should().BeTrue();

        using var doc = WordprocessingDocument.Open(outputPath, false);
        var body = doc.MainDocumentPart!.Document.Body!;

        var tables = body.Elements<Table>().ToList();
        tables.Should().NotBeEmpty();

        var rows = tables[0].Elements<TableRow>().ToList();
        rows.Should().HaveCount(3); // header + 2 data rows
    }

    [Fact]
    public void Process_ImagePlaceholder_InsertsImage()
    {
        var templatePath = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "Logo");

        // Minimal 1x1 PNG
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
        var imagePath = Path.Combine(_tempDir, "logo.png");
        File.WriteAllBytes(imagePath, pngBytes);

        var configJson = $$"""
        {
          "placeholders": {
            "Logo": {
              "type": "image",
              "source": "logo.png",
              "widthCm": 5.0,
              "heightCm": 3.0
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var outputPath = Path.Combine(_tempDir, "output_image.docx");
        var config = TemplateConfig.Load(configPath);
        var processor = new TemplateProcessor(verbose: false);

        processor.Process(templatePath, config, outputPath);

        File.Exists(outputPath).Should().BeTrue();

        using var doc = WordprocessingDocument.Open(outputPath, false);
        var drawings = doc.MainDocumentPart!.Document.Body!
            .Descendants<Drawing>().ToList();
        drawings.Should().NotBeEmpty();
    }

    [Fact]
    public void Process_FileObjectPlaceholder_EmbedsFile()
    {
        var templatePath = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "Attachment");

        var txtPath = Path.Combine(_tempDir, "report.txt");
        File.WriteAllText(txtPath, "Sample report content");

        var configJson = $$"""
        {
          "placeholders": {
            "Attachment": {
              "type": "file",
              "source": "report.txt",
              "displayName": "Monthly Report"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var outputPath = Path.Combine(_tempDir, "output_file.docx");
        var config = TemplateConfig.Load(configPath);
        var processor = new TemplateProcessor(verbose: false);

        processor.Process(templatePath, config, outputPath);

        File.Exists(outputPath).Should().BeTrue();

        using var doc = WordprocessingDocument.Open(outputPath, false);
        var oleParts = doc.MainDocumentPart!.EmbeddedObjectParts.ToList();
        oleParts.Should().ContainSingle();

        byte[] partBytes;
        using (var partStream = oleParts[0].GetStream())
        using (var memory = new MemoryStream())
        {
            partStream.CopyTo(memory);
            partBytes = memory.ToArray();
        }

        partBytes.Take(8).Should().Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });

        using (var cfbf = new MemoryStream(partBytes))
        using (var root = RootStorage.Open(cfbf))
        {
            root.CLSID.Should().Be(new Guid("0003000C-0000-0000-C000-000000000046"));

            var entryNames = root.EnumerateEntries()
                .Where(e => e.Type == EntryType.Stream)
                .Select(e => e.Name)
                .ToList();
            entryNames.Should().BeEquivalentTo(new[] { "CompObj", "ObjInfo", "Ole10Native" });

            using var oleStream = root.OpenStream("Ole10Native");
            var nativeBuf = new byte[oleStream.Length];
            oleStream.Read(nativeBuf, 0, nativeBuf.Length);

            Encoding.GetEncoding(1252).GetString(nativeBuf).Should().Contain("report.txt");
            Encoding.ASCII.GetString(nativeBuf).Should().Contain("Sample report content");
        }

        doc.MainDocumentPart!.ImageParts
            .Should().ContainSingle(p => p.ContentType == "image/png");
    }

    [Fact]
    public void Process_FileObjectPlaceholder_OfficeFile_StoresRawPackage()
    {
        var templatePath = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "Attachment");

        var nestedSource = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "Inside");
        var nestedDocx = Path.Combine(_tempDir, "nested.docx");
        File.Move(nestedSource, nestedDocx);

        var configJson = $$"""
        {
          "placeholders": {
            "Attachment": {
              "type": "file",
              "source": "nested.docx",
              "displayName": "Nested Doc"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var outputPath = Path.Combine(_tempDir, "output_office.docx");
        var config = TemplateConfig.Load(configPath);
        var processor = new TemplateProcessor(verbose: false);

        processor.Process(templatePath, config, outputPath);

        using var doc = WordprocessingDocument.Open(outputPath, false);
        var pkgParts = doc.MainDocumentPart!.EmbeddedPackageParts.ToList();
        pkgParts.Should().ContainSingle();

        byte[] partBytes;
        using (var partStream = pkgParts[0].GetStream())
        using (var memory = new MemoryStream())
        {
            partStream.CopyTo(memory);
            partBytes = memory.ToArray();
        }

        partBytes.Take(4).Should().Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 });
    }

    [Fact]
    public void DryRun_ValidConfig_ReturnsValid()
    {
        var templatePath = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "Content");
        var mdPath = Path.Combine(_tempDir, "content.md");
        File.WriteAllText(mdPath, "# Test");

        var configJson = $$"""
        {
          "placeholders": {
            "Content": {
              "type": "markdown",
              "source": "content.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var config = TemplateConfig.Load(configPath);
        var processor = new TemplateProcessor(verbose: false);

        var result = processor.DryRun(templatePath, config);

        result.IsValid.Should().BeTrue();
        result.PlaceholdersFound.Should().Contain("Content");
    }

    [Fact]
    public void Process_SplitRunPlaceholder_ReplacesCorrectly()
    {
        var templatePath = TestDocxHelper.CreateTemplateWithSplitRuns(_tempDir, "Content");
        var mdPath = Path.Combine(_tempDir, "content.md");
        File.WriteAllText(mdPath, "Replaced successfully.");

        var configJson = $$"""
        {
          "placeholders": {
            "Content": {
              "type": "markdown",
              "source": "content.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var outputPath = Path.Combine(_tempDir, "output_split.docx");
        var config = TemplateConfig.Load(configPath);
        var processor = new TemplateProcessor(verbose: false);

        processor.Process(templatePath, config, outputPath);

        using var doc = WordprocessingDocument.Open(outputPath, false);
        var body = doc.MainDocumentPart!.Document.Body!;

        body.InnerText.Should().Contain("Replaced successfully");
        body.InnerText.Should().NotContain("{{");
    }
}
