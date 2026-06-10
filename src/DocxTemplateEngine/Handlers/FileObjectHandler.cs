using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Engine;
using DocxTemplateEngine.Models;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;

namespace DocxTemplateEngine.Handlers;

/// <summary>
/// Embeds a file as an OLE object (clickable icon) in the DOCX document.
/// Office formats are stored as raw OOXML packages; everything else is
/// wrapped in a CFBF Object Packager container so Word can activate the
/// embed regardless of the host's installed apps.
/// </summary>
public class FileObjectHandler : IPlaceholderHandler
{
    private static readonly HashSet<string> OfficeExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"
    };

    private static readonly Dictionary<string, string> OfficeProgIdMap = new(StringComparer.OrdinalIgnoreCase)
    {
        [".xlsx"] = "Excel.Sheet.12",
        [".xls"] = "Excel.Sheet.8",
        [".docx"] = "Word.Document.12",
        [".doc"] = "Word.Document.8",
        [".pptx"] = "PowerPoint.Show.12",
        [".ppt"] = "PowerPoint.Show.8",
    };

    private static readonly Dictionary<string, string> OfficeContentTypeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        [".xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        [".xls"] = "application/vnd.ms-excel",
        [".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        [".doc"] = "application/msword",
        [".pptx"] = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        [".ppt"] = "application/vnd.ms-powerpoint",
    };

    private const string OleObjectContentType = "application/vnd.openxmlformats-officedocument.oleObject";
    private const string GenericProgId = "Package";

    public void Replace(PlaceholderMatch match, string source, WordprocessingDocument document, PlaceholderEntry entry)
    {
        var mainPart = document.MainDocumentPart!;
        var ext = System.IO.Path.GetExtension(source).ToLowerInvariant();
        var displayName = entry.DisplayName ?? System.IO.Path.GetFileName(source);
        var fileName = System.IO.Path.GetFileName(source);

        string embeddedPartId;
        string progId;

        if (OfficeExtensions.Contains(ext))
        {
            var contentType = OfficeContentTypeMap[ext];
            var part = mainPart.AddEmbeddedPackagePart(contentType);
            using (var fileStream = System.IO.File.OpenRead(source))
            {
                part.FeedData(fileStream);
            }
            embeddedPartId = mainPart.GetIdOfPart(part);
            progId = OfficeProgIdMap[ext];
        }
        else
        {
            var part = mainPart.AddEmbeddedObjectPart(OleObjectContentType);
            var cfbfBytes = OlePackageBuilder.Build(source, fileName);
            using (var ms = new MemoryStream(cfbfBytes, writable: false))
            {
                part.FeedData(ms);
            }
            embeddedPartId = mainPart.GetIdOfPart(part);
            progId = GenericProgId;
        }

        var iconImagePart = mainPart.AddImagePart(ImagePartType.Png);
        var iconImageId = mainPart.GetIdOfPart(iconImagePart);
        iconImagePart.FeedData(new MemoryStream(IconBytes, writable: false));

        var oleObject = CreateOleObjectParagraph(embeddedPartId, iconImageId, progId, displayName);

        PlaceholderFinder.RemovePlaceholderText(match);

        var paragraph = match.Paragraph;
        var parent = paragraph.Parent!;

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
        var oleObj = new DocumentFormat.OpenXml.Wordprocessing.EmbeddedObject();

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
