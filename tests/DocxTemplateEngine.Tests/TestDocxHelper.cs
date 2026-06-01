using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplateEngine.Tests;

/// <summary>
/// Helper to create DOCX files for testing.
/// </summary>
public static class TestDocxHelper
{
    public static string CreateTemplateWithPlaceholders(string dir, params string[] placeholderNames)
    {
        var path = Path.Combine(dir, $"template_{Guid.NewGuid():N}.docx");

        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        foreach (var name in placeholderNames)
        {
            var paragraph = new Paragraph(
                new Run(new Text($"{{{{{name}}}}}") { Space = SpaceProcessingModeValues.Preserve })
            );
            mainPart.Document.Body!.Append(paragraph);
        }

        mainPart.Document.Save();
        return path;
    }

    public static string CreateTemplateWithSplitRuns(string dir, string placeholderName)
    {
        var path = Path.Combine(dir, $"split_template_{Guid.NewGuid():N}.docx");

        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        // Simulate Word splitting the placeholder across multiple runs
        var paragraph = new Paragraph(
            new Run(new Text("{{") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new Text(placeholderName) { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new Text("}}") { Space = SpaceProcessingModeValues.Preserve })
        );
        mainPart.Document.Body!.Append(paragraph);
        mainPart.Document.Save();
        return path;
    }

    public static string CreateTemplateWithMixedContent(string dir, string placeholderName)
    {
        var path = Path.Combine(dir, $"mixed_template_{Guid.NewGuid():N}.docx");

        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var paragraph = new Paragraph(
            new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new Text($"{{{{{placeholderName}}}}}") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new Text(" After") { Space = SpaceProcessingModeValues.Preserve })
        );
        mainPart.Document.Body!.Append(paragraph);
        mainPart.Document.Save();
        return path;
    }
}
