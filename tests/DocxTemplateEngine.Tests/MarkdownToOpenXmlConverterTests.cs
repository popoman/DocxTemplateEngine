using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Converters;
using FluentAssertions;

namespace DocxTemplateEngine.Tests;

public class MarkdownToOpenXmlConverterTests : IDisposable
{
    private readonly string _tempDir;

    public MarkdownToOpenXmlConverterTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"DocxTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    private (WordprocessingDocument Doc, MarkdownToOpenXmlConverter Converter) CreateDocAndConverter()
    {
        var path = Path.Combine(_tempDir, $"test_{Guid.NewGuid():N}.docx");
        var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        return (doc, new MarkdownToOpenXmlConverter(doc));
    }

    [Fact]
    public void Convert_SimpleParagraph_ReturnsParagraph()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var elements = converter.Convert("Hello, World!");
            elements.Should().HaveCount(1);
            elements[0].Should().BeOfType<Paragraph>();
            elements[0].InnerText.Should().Be("Hello, World!");
        }
    }

    [Fact]
    public void Convert_Heading_ReturnsParagraphWithHeadingStyle()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var elements = converter.Convert("# My Heading");
            elements.Should().NotBeEmpty();

            var para = elements.OfType<Paragraph>()
                .FirstOrDefault(p => p.GetFirstChild<ParagraphProperties>()?.ParagraphStyleId != null);
            para.Should().NotBeNull();

            var style = para!.GetFirstChild<ParagraphProperties>()?.ParagraphStyleId;
            style.Should().NotBeNull();
            style!.Val!.Value.Should().Be("Heading1");
        }
    }

    [Fact]
    public void Convert_BoldText_HasBoldProperty()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var elements = converter.Convert("**bold text**");
            elements.Should().HaveCount(1);

            var runs = elements[0].Descendants<Run>().ToList();
            runs.Should().NotBeEmpty();
            runs.Any(r => r.RunProperties?.Bold != null).Should().BeTrue();
        }
    }

    [Fact]
    public void Convert_ItalicText_HasItalicProperty()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var elements = converter.Convert("*italic text*");
            elements.Should().HaveCount(1);

            var runs = elements[0].Descendants<Run>().ToList();
            runs.Should().NotBeEmpty();
            runs.Any(r => r.RunProperties?.Italic != null).Should().BeTrue();
        }
    }

    [Fact]
    public void Convert_InlineCode_HasMonospaceFont()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var elements = converter.Convert("Use `Console.WriteLine`");
            elements.Should().HaveCount(1);

            var runs = elements[0].Descendants<Run>().ToList();
            runs.Any(r => r.RunProperties?.RunFonts?.Ascii?.Value == "Courier New")
                .Should().BeTrue();
        }
    }

    [Fact]
    public void Convert_CodeBlock_HasShadingAndMonospace()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var elements = converter.Convert("```\nvar x = 1;\n```");
            elements.Should().NotBeEmpty();

            var para = elements[0] as Paragraph;
            para.Should().NotBeNull();

            var shading = para!.GetFirstChild<ParagraphProperties>()?.Shading;
            shading.Should().NotBeNull();
        }
    }

    [Fact]
    public void Convert_UnorderedList_HasNumberingProperties()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var elements = converter.Convert("- Item 1\n- Item 2\n- Item 3");
            elements.Should().HaveCountGreaterThanOrEqualTo(3);

            foreach (var el in elements)
            {
                if (el is Paragraph p)
                {
                    var numProps = p.GetFirstChild<ParagraphProperties>()?
                        .GetFirstChild<NumberingProperties>();
                    numProps.Should().NotBeNull();
                }
            }
        }
    }

    [Fact]
    public void Convert_MultipleParagraphs_ReturnsMultipleElements()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var elements = converter.Convert("First paragraph.\n\nSecond paragraph.");
            elements.Should().HaveCount(2);
        }
    }

    [Fact]
    public void Convert_LinkReferenceDefinitions_NotRenderedAsText()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var markdown = "Click [here][link1] for details.\n\n[link1]: https://example.com";
            var elements = converter.Convert(markdown);

            // Should not contain the type name "LinkReferenceDefinitionGroup"
            var fullText = string.Join(" ", elements.Select(e => e.InnerText));
            fullText.Should().NotContain("LinkReferenceDefinition");
            fullText.Should().Contain("here");
        }
    }

    [Fact]
    public void Convert_PipeTable_ReturnsTableElement()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var markdown = "| Name | Age |\n| --- | --- |\n| Alice | 30 |\n| Bob | 25 |";
            var elements = converter.Convert(markdown);

            elements.Should().ContainSingle();
            elements[0].Should().BeOfType<Table>();

            var table = (Table)elements[0];
            var rows = table.Elements<TableRow>().ToList();
            rows.Should().HaveCount(3); // header + 2 data rows
        }
    }

    [Fact]
    public void Convert_PipeTable_HeaderCellsHaveNoExplicitRunProperties()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var markdown = "| Name | Age |\n| --- | --- |\n| Alice | 30 |";
            var elements = converter.Convert(markdown);

            var table = (Table)elements[0];
            var headerRow = table.Elements<TableRow>().First();
            var firstCell = headerRow.Elements<TableCell>().First();
            var run = firstCell.Descendants<Run>().First();
            run.RunProperties.Should().BeNull();
        }
    }

    [Fact]
    public void Convert_PipeTable_CellTextExtractedCorrectly()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var markdown = "| Name | Value |\n| --- | --- |\n| **bold** | `code` |";
            var elements = converter.Convert(markdown);

            var table = (Table)elements[0];
            var dataRow = table.Elements<TableRow>().Last();
            var cells = dataRow.Elements<TableCell>().ToList();

            // Verify cell text is actual content, not type names
            cells[0].InnerText.Should().Be("bold");
            cells[1].InnerText.Should().Be("code");
        }
    }

    [Fact]
    public void Convert_MarkdownWithTextAndTable_ReturnsBothElements()
    {
        var (doc, converter) = CreateDocAndConverter();
        using (doc)
        {
            var markdown = "# Header\n\nSome text.\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\nMore text.";
            var elements = converter.Convert(markdown);

            // Should contain: heading, paragraph, table, paragraph
            elements.Should().HaveCount(4);
            elements[0].Should().BeOfType<Paragraph>(); // heading
            elements[1].Should().BeOfType<Paragraph>(); // "Some text."
            elements[2].Should().BeOfType<Table>();       // the table
            elements[3].Should().BeOfType<Paragraph>(); // "More text."
        }
    }
}
