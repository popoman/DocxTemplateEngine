using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Engine;
using DocxTemplateEngine.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxTemplateEngine.Handlers;

public class ImageHandler : IPlaceholderHandler
{
    private const long EmuPerCm = 360000;

    public void Replace(PlaceholderMatch match, string source, WordprocessingDocument document, PlaceholderEntry entry)
    {
        var mainPart = document.MainDocumentPart!;
        var imagePart = AddImagePart(mainPart, source);
        var relationshipId = mainPart.GetIdOfPart(imagePart);

        var (widthEmu, heightEmu) = GetImageDimensions(source, entry);

        var drawingElement = CreateImageDrawing(relationshipId, widthEmu, heightEmu, Path.GetFileName(source));

        // Remove placeholder text and insert image
        var anchorRun = PlaceholderFinder.RemovePlaceholderText(match);

        // Clear remaining text in the anchor run and add the drawing
        var textEl = anchorRun.GetFirstChild<Text>();
        if (textEl != null && string.IsNullOrEmpty(textEl.Text))
            textEl.Remove();

        anchorRun.Append(drawingElement);
    }

    private static ImagePart AddImagePart(MainDocumentPart mainPart, string imagePath)
    {
        var ext = Path.GetExtension(imagePath).ToLowerInvariant();
        var imageType = ext switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".bmp" => ImagePartType.Bmp,
            ".gif" => ImagePartType.Gif,
            _ => throw new InvalidOperationException($"Unsupported image format: {ext}")
        };

        var imagePart = mainPart.AddImagePart(imageType);
        using var stream = File.OpenRead(imagePath);
        imagePart.FeedData(stream);
        return imagePart;
    }

    private static (long Width, long Height) GetImageDimensions(string imagePath, PlaceholderEntry entry)
    {
        // Use configured dimensions if provided
        if (entry.WidthCm.HasValue && entry.HeightCm.HasValue)
        {
            return (
                (long)(entry.WidthCm.Value * EmuPerCm),
                (long)(entry.HeightCm.Value * EmuPerCm)
            );
        }

        // Read image dimensions from binary headers (no System.Drawing dependency)
        var (pixelWidth, pixelHeight) = ReadImageDimensions(imagePath);

        // Assume 96 DPI if we can't determine resolution
        const double defaultDpi = 96.0;
        var widthInches = pixelWidth / defaultDpi;
        var heightInches = pixelHeight / defaultDpi;

        const long emuPerInch = 914400;
        var widthEmu = (long)(widthInches * emuPerInch);
        var heightEmu = (long)(heightInches * emuPerInch);

        // Constrain to max 16cm width (standard page width minus margins)
        const long maxWidth = 16 * EmuPerCm;
        if (widthEmu > maxWidth)
        {
            var ratio = (double)maxWidth / widthEmu;
            widthEmu = maxWidth;
            heightEmu = (long)(heightEmu * ratio);
        }

        return (widthEmu, heightEmu);
    }

    private static (int Width, int Height) ReadImageDimensions(string imagePath)
    {
        var ext = Path.GetExtension(imagePath).ToLowerInvariant();
        var bytes = File.ReadAllBytes(imagePath);

        if (ext == ".png")
        {
            // PNG: width at offset 16 (4 bytes BE), height at offset 20 (4 bytes BE)
            if (bytes.Length > 24)
            {
                var w = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
                var h = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
                return (w, h);
            }
        }
        else if (ext is ".jpg" or ".jpeg")
        {
            // JPEG: search for SOF0 marker (0xFF 0xC0)
            for (int i = 0; i < bytes.Length - 9; i++)
            {
                if (bytes[i] == 0xFF && (bytes[i + 1] == 0xC0 || bytes[i + 1] == 0xC2))
                {
                    var h = (bytes[i + 5] << 8) | bytes[i + 6];
                    var w = (bytes[i + 7] << 8) | bytes[i + 8];
                    return (w, h);
                }
            }
        }

        // Fallback default
        return (400, 300);
    }

    private static Drawing CreateImageDrawing(string relationshipId, long widthEmu, long heightEmu, string fileName)
    {
        return new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = (uint)1, Name = fileName },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0, Name = fileName },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle }))
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = 0,
                DistanceFromBottom = 0,
                DistanceFromLeft = 0,
                DistanceFromRight = 0
            });
    }
}
