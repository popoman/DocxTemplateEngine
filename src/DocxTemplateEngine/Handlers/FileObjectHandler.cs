using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Engine;
using DocxTemplateEngine.Models;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;

namespace DocxTemplateEngine.Handlers;

/// <summary>
/// Embeds a file as an OLE object (embedded package) in the DOCX document.
/// The file appears as a clickable icon with a display name.
/// </summary>
public class FileObjectHandler : IPlaceholderHandler
{
    private static readonly Dictionary<string, string> ProgIdMap = new(StringComparer.OrdinalIgnoreCase)
    {
        [".xlsx"] = "Excel.Sheet.12",
        [".xls"] = "Excel.Sheet.8",
        [".docx"] = "Word.Document.12",
        [".doc"] = "Word.Document.8",
        [".pptx"] = "PowerPoint.Show.12",
        [".ppt"] = "PowerPoint.Show.8",
        [".pdf"] = "AcroExch.Document.DC",
        [".txt"] = "Package",
        [".csv"] = "Excel.Sheet.12",
        [".zip"] = "Package",
    };

    private static readonly Dictionary<string, string> ContentTypeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        [".xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        [".xls"] = "application/vnd.ms-excel",
        [".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        [".doc"] = "application/msword",
        [".pptx"] = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        [".ppt"] = "application/vnd.ms-powerpoint",
        [".pdf"] = "application/pdf",
        [".txt"] = "text/plain",
        [".csv"] = "text/csv",
        [".zip"] = "application/zip",
    };

    public void Replace(PlaceholderMatch match, string source, WordprocessingDocument document, PlaceholderEntry entry)
    {
        var mainPart = document.MainDocumentPart!;
        var ext = System.IO.Path.GetExtension(source).ToLowerInvariant();
        var displayName = entry.DisplayName ?? System.IO.Path.GetFileName(source);

        // Embed the file as a package part
        var contentType = ContentTypeMap.GetValueOrDefault(ext, "application/octet-stream");
        var embeddedPart = mainPart.AddEmbeddedPackagePart(contentType);

        using (var fileStream = System.IO.File.OpenRead(source))
        {
            embeddedPart.FeedData(fileStream);
        }

        var embeddedPartId = mainPart.GetIdOfPart(embeddedPart);

        // Create an icon image for the embedded object using the bundled static PNG.
        var iconImagePart = mainPart.AddImagePart(ImagePartType.Png);
        var iconImageId = mainPart.GetIdOfPart(iconImagePart);
        iconImagePart.FeedData(new MemoryStream(IconBytes, writable: false));

        var progId = ProgIdMap.GetValueOrDefault(ext, "Package");

        // Build the OLE object paragraph
        var oleObject = CreateOleObjectParagraph(embeddedPartId, iconImageId, progId, displayName);

        // Remove placeholder and insert the OLE object
        PlaceholderFinder.RemovePlaceholderText(match);

        var paragraph = match.Paragraph;
        var parent = paragraph.Parent!;

        // Check if paragraph is now empty
        var remainingText = paragraph.InnerText.Trim();
        if (string.IsNullOrEmpty(remainingText))
        {
            parent.InsertAfter(oleObject, paragraph);
            paragraph.Remove();
        }
        else
        {
            parent.InsertAfter(oleObject, paragraph);
        }
    }

    private static Paragraph CreateOleObjectParagraph(
        string embeddedPartId, string iconImageId, string progId, string displayName)
    {
        var paragraph = new Paragraph();

        var run = new Run();

        // Create the OLE object element
        var oleObj = new DocumentFormat.OpenXml.Wordprocessing.EmbeddedObject();

        // VML Shape for the icon representation
        var shape = new Shape
        {
            Id = $"_x0000_i{Guid.NewGuid():N}".Substring(0, 20),
            Style = "width:64pt;height:64pt",
            Type = "#_x0000_t75"
        };

        var imageData = new ImageData
        {
            RelationshipId = iconImageId,
            Title = displayName
        };
        shape.Append(imageData);

        oleObj.Append(shape);

        // OLE Object reference
        var oleObjElement = new DocumentFormat.OpenXml.Vml.Office.OleObject
        {
            Type = OleValues.Embed,
            ProgId = progId,
            ShapeId = shape.Id,
            DrawAspect = OleDrawAspectValues.Icon,
            ObjectId = $"_ole_{Guid.NewGuid():N}".Substring(0, 20),
            Id = embeddedPartId
        };

        oleObj.Append(oleObjElement);

        run.Append(oleObj);
        paragraph.Append(run);

        // Add a run with the display name after the object
        var nameRun = new Run();
        nameRun.Append(new Text($" {displayName}") { Space = SpaceProcessingModeValues.Preserve });
        paragraph.Append(nameRun);

        return paragraph;
    }

    private const string IconResourceName = "DocxTemplateEngine.Resources.file-icon.png";

    private static readonly byte[] IconBytes = LoadIconBytes();

    private static byte[] LoadIconBytes()
    {
        using var stream = typeof(FileObjectHandler).Assembly.GetManifestResourceStream(IconResourceName)
            ?? throw new InvalidOperationException(
                $"Embedded icon resource '{IconResourceName}' not found in assembly.");
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }
}
