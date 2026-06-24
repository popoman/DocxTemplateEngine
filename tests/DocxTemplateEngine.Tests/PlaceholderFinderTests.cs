using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Engine;
using FluentAssertions;

namespace DocxTemplateEngine.Tests;

public class PlaceholderFinderTests : IDisposable
{
    private readonly string _tempDir;

    public PlaceholderFinderTests()
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
    public void FindAll_SinglePlaceholder_ReturnsMatch()
    {
        var path = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "TestName");
        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart!.Document.Body!;

        var matches = PlaceholderFinder.FindAll(body);

        matches.Should().HaveCount(1);
        matches[0].Name.Should().Be("TestName");
    }

    [Fact]
    public void FindAll_MultiplePlaceholders_ReturnsAll()
    {
        var path = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "First", "Second", "Third");
        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart!.Document.Body!;

        var matches = PlaceholderFinder.FindAll(body);

        matches.Should().HaveCount(3);
        matches.Select(m => m.Name).Should().Contain(["First", "Second", "Third"]);
    }

    [Fact]
    public void FindAll_SplitRunPlaceholder_ReturnsMatch()
    {
        var path = TestDocxHelper.CreateTemplateWithSplitRuns(_tempDir, "SplitTest");
        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart!.Document.Body!;

        var matches = PlaceholderFinder.FindAll(body);

        matches.Should().HaveCount(1);
        matches[0].Name.Should().Be("SplitTest");
        matches[0].Runs.Should().HaveCount(3);
    }

    [Fact]
    public void FindAll_PlaceholderWithEmptyRunBetweenTokens_ReturnsMatch()
    {
        var path = Path.Combine(_tempDir, "split_with_empty_run.docx");
        using (var doc = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("{{") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new RunProperties(new Bold())),
                    new Run(new Text("SplitTest") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new Text("}}") { Space = SpaceProcessingModeValues.Preserve })
                )
            ));
            mainPart.Document.Save();
        }

        using var readDoc = WordprocessingDocument.Open(path, true);
        var body = readDoc.MainDocumentPart!.Document.Body!;

        var matches = PlaceholderFinder.FindAll(body);

        matches.Should().HaveCount(1);
        matches[0].Name.Should().Be("SplitTest");

        PlaceholderFinder.RemovePlaceholderText(matches[0]);
        PlaceholderFinder.FindAll(body).Should().BeEmpty();
    }

    [Fact]
    public void FindAll_NoPlaceholders_ReturnsEmpty()
    {
        var path = Path.Combine(_tempDir, "empty.docx");
        using var doc = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(
            new Paragraph(new Run(new Text("No placeholders here.")))
        ));
        mainPart.Document.Save();
        doc.Dispose();

        using var readDoc = WordprocessingDocument.Open(path, false);
        var body = readDoc.MainDocumentPart!.Document.Body!;

        var matches = PlaceholderFinder.FindAll(body);
        matches.Should().BeEmpty();
    }

    [Fact]
    public void RemovePlaceholderText_ClearsPlaceholder()
    {
        var path = TestDocxHelper.CreateTemplateWithPlaceholders(_tempDir, "ToRemove");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var matches = PlaceholderFinder.FindAll(body);
        matches.Should().HaveCount(1);

        var run = PlaceholderFinder.RemovePlaceholderText(matches[0]);
        run.Should().NotBeNull();

        // After removal, the placeholder text should be gone
        var newMatches = PlaceholderFinder.FindAll(body);
        newMatches.Should().BeEmpty();
    }
}
