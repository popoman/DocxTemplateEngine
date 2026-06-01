using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Converters;
using DocxTemplateEngine.Engine;
using DocxTemplateEngine.Models;
using FluentAssertions;

namespace DocxTemplateEngine.Tests;

public class ExistingTablePopulationTests : IDisposable
{
    private readonly string _tempDir;

    public ExistingTablePopulationTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"DocxTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    /// <summary>
    /// Creates a DOCX template with an existing table: header row + a row containing {{placeholder}} in the first cell.
    /// </summary>
    private string CreateTemplateWithExistingTable(string placeholderName, string[] headers)
    {
        var path = Path.Combine(_tempDir, $"table_template_{Guid.NewGuid():N}.docx");

        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var table = new Table();

        // Table properties with borders
        var tblProps = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" }
            ),
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
        );
        table.Append(tblProps);

        // Header row
        var headerRow = new TableRow();
        foreach (var header in headers)
        {
            var cell = new TableCell();
            var cellProps = new TableCellProperties(
                new TableCellWidth { Width = "2400", Type = TableWidthUnitValues.Dxa },
                new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9E2F3" }
            );
            cell.Append(cellProps);

            var run = new Run();
            run.Append(new RunProperties(new Bold()));
            run.Append(new Text(header) { Space = SpaceProcessingModeValues.Preserve });
            cell.Append(new Paragraph(run));
            headerRow.Append(cell);
        }
        table.Append(headerRow);

        // Placeholder row — {{placeholder}} in first cell, empty cells for the rest
        var placeholderRow = new TableRow();
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = new TableCell();
            var cellProps = new TableCellProperties(
                new TableCellWidth { Width = "2400", Type = TableWidthUnitValues.Dxa }
            );
            cell.Append(cellProps);

            if (i == 0)
            {
                cell.Append(new Paragraph(
                    new Run(new Text($"{{{{{placeholderName}}}}}") { Space = SpaceProcessingModeValues.Preserve })
                ));
            }
            else
            {
                cell.Append(new Paragraph());
            }

            placeholderRow.Append(cell);
        }
        table.Append(placeholderRow);

        mainPart.Document.Body!.Append(table);
        mainPart.Document.Save();
        return path;
    }

    [Fact]
    public void PopulateExistingTable_AppendsDataRowsAndRemovesPlaceholderRow()
    {
        var templatePath = CreateTemplateWithExistingTable("DataTable", ["Name", "Score"]);

        var mdPath = Path.Combine(_tempDir, "data.md");
        File.WriteAllText(mdPath, "| Name | Score |\n|------|-------|\n| Alice | 95 |\n| Bob | 87 |\n| Carol | 91 |");

        var configJson = """
        {
          "placeholders": {
            "DataTable": {
              "type": "markdownTable",
              "source": "data.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var outputPath = Path.Combine(_tempDir, "output_populate.docx");
        var config = TemplateConfig.Load(configPath);
        var processor = new TemplateProcessor(verbose: false);

        processor.Process(templatePath, config, outputPath);

        File.Exists(outputPath).Should().BeTrue();

        using var doc = WordprocessingDocument.Open(outputPath, false);
        var body = doc.MainDocumentPart!.Document.Body!;

        var tables = body.Elements<Table>().ToList();
        tables.Should().HaveCount(1, "should reuse the existing table, not create a new one");

        var rows = tables[0].Elements<TableRow>().ToList();
        rows.Should().HaveCount(4, "header + 3 data rows (placeholder row removed)");

        // Verify header row is still styled
        var headerCells = rows[0].Elements<TableCell>().ToList();
        headerCells[0].InnerText.Should().Be("Name");
        headerCells[1].InnerText.Should().Be("Score");
        var headerShading = headerCells[0].GetFirstChild<TableCellProperties>()?.GetFirstChild<Shading>();
        headerShading.Should().NotBeNull("header styling should be preserved");

        // Verify data rows
        rows[1].Elements<TableCell>().First().InnerText.Should().Be("Alice");
        rows[2].Elements<TableCell>().First().InnerText.Should().Be("Bob");
        rows[3].Elements<TableCell>().First().InnerText.Should().Be("Carol");

        // Verify no placeholder text remains
        body.InnerText.Should().NotContain("{{DataTable}}");
    }

    [Fact]
    public void PopulateExistingTable_PreservesCellWidths()
    {
        var templatePath = CreateTemplateWithExistingTable("Data", ["Col1", "Col2"]);

        var mdPath = Path.Combine(_tempDir, "data.md");
        File.WriteAllText(mdPath, "| Col1 | Col2 |\n|------|------|\n| A | B |");

        var configJson = """
        {
          "placeholders": {
            "Data": {
              "type": "markdownTable",
              "source": "data.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var outputPath = Path.Combine(_tempDir, "output_widths.docx");
        var config = TemplateConfig.Load(configPath);
        new TemplateProcessor().Process(templatePath, config, outputPath);

        using var doc = WordprocessingDocument.Open(outputPath, false);
        var table = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var dataRow = table.Elements<TableRow>().Last();
        var cellWidth = dataRow.Elements<TableCell>().First()
            .GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellWidth>();

        cellWidth.Should().NotBeNull("cell widths should be cloned from template row");
        cellWidth!.Width!.Value.Should().Be("2400");
    }

    [Fact]
    public void PopulateExistingTable_ThreeColumns()
    {
        var templatePath = CreateTemplateWithExistingTable("Info", ["Name", "Age", "City"]);

        var mdPath = Path.Combine(_tempDir, "info.md");
        File.WriteAllText(mdPath, "| Name | Age | City |\n|------|-----|------|\n| Alice | 30 | NYC |\n| Bob | 25 | LA |");

        var configJson = """
        {
          "placeholders": {
            "Info": {
              "type": "markdownTable",
              "source": "info.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var outputPath = Path.Combine(_tempDir, "output_3col.docx");
        var config = TemplateConfig.Load(configPath);
        new TemplateProcessor().Process(templatePath, config, outputPath);

        using var doc = WordprocessingDocument.Open(outputPath, false);
        var table = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var rows = table.Elements<TableRow>().ToList();

        rows.Should().HaveCount(3, "header + 2 data rows");

        // Verify each data row has 3 cells
        foreach (var row in rows.Skip(1))
        {
            row.Elements<TableCell>().Should().HaveCount(3);
        }

        // Check data content
        var lastRow = rows[2];
        var cells = lastRow.Elements<TableCell>().ToList();
        cells[0].InnerText.Should().Be("Bob");
        cells[1].InnerText.Should().Be("25");
        cells[2].InnerText.Should().Be("LA");
    }

    [Fact]
    public void ExtractDataRows_SkipsHeaderReturnsDataOnly()
    {
        var path = Path.Combine(_tempDir, "test.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        doc.AddMainDocumentPart().Document = new Document(new Body());

        var converter = new MarkdownTableConverter(doc);
        var rows = converter.ExtractDataRows("| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |");

        rows.Should().HaveCount(2);
        rows[0].Should().BeEquivalentTo(["1", "2"]);
        rows[1].Should().BeEquivalentTo(["3", "4"]);
    }

    [Fact]
    public void ExtractDataRows_NoTable_ReturnsEmpty()
    {
        var path = Path.Combine(_tempDir, "test.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        doc.AddMainDocumentPart().Document = new Document(new Body());

        var converter = new MarkdownTableConverter(doc);
        var rows = converter.ExtractDataRows("Just plain text.");

        rows.Should().BeEmpty();
    }
}
