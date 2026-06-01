using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Converters;
using FluentAssertions;

namespace DocxTemplateEngine.Tests;

public class MarkdownTableConverterTests : IDisposable
{
    private readonly string _tempDir;

    public MarkdownTableConverterTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"DocxTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    private (WordprocessingDocument Doc, MarkdownTableConverter Converter) CreateDocAndConverter()
    {
        var path = Path.Combine(_tempDir, $"test_{Guid.NewGuid():N}.docx");
        var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        return (doc, new MarkdownTableConverter(doc));
    }

    [Fact]
    public void Convert_SimpleTable_ReturnsTable()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var markdown = "| Name | Age |\n|------|-----|\n| Alice | 30 |\n| Bob | 25 |";
            var table = converter.Convert(markdown);

            table.Should().NotBeNull();

            var rows = table!.Elements<TableRow>().ToList();
            rows.Should().HaveCount(3); // header + 2 data rows
        }
    }

    [Fact]
    public void Convert_HeaderRow_HasBoldAndShading()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var markdown = "| Col1 | Col2 |\n|------|------|\n| A | B |";
            var table = converter.Convert(markdown);

            var firstRow = table!.Elements<TableRow>().First();

            // Check table header property
            var trProps = firstRow.GetFirstChild<TableRowProperties>();
            trProps.Should().NotBeNull();
            trProps!.GetFirstChild<TableHeader>().Should().NotBeNull();

            // Check cell shading
            var firstCell = firstRow.Elements<TableCell>().First();
            var shading = firstCell.GetFirstChild<TableCellProperties>()?.GetFirstChild<Shading>();
            shading.Should().NotBeNull();
        }
    }

    [Fact]
    public void Convert_NoTable_ReturnsNull()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var table = converter.Convert("Just plain text, no table here.");
            table.Should().BeNull();
        }
    }

    [Fact]
    public void Convert_Table_HasBorders()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var markdown = "| A | B |\n|---|---|\n| 1 | 2 |";
            var table = converter.Convert(markdown);

            var tblProps = table!.GetFirstChild<TableProperties>();
            tblProps.Should().NotBeNull();

            var borders = tblProps!.GetFirstChild<TableBorders>();
            borders.Should().NotBeNull();
        }
    }
}
