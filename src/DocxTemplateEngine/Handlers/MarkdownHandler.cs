using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Converters;
using DocxTemplateEngine.Engine;
using DocxTemplateEngine.Models;

namespace DocxTemplateEngine.Handlers;

public class MarkdownHandler : IPlaceholderHandler
{
    public void Replace(PlaceholderMatch match, string source, WordprocessingDocument document, PlaceholderEntry entry)
    {
        var markdown = File.ReadAllText(source);
        var converter = new MarkdownToOpenXmlConverter(document);
        var elements = converter.Convert(markdown);

        if (elements.Count == 0) return;

        var paragraph = match.Paragraph;

        // Remove the placeholder text from runs
        PlaceholderFinder.RemovePlaceholderText(match);

        // Check if the paragraph has any remaining text
        var remainingText = paragraph.InnerText.Trim();
        if (string.IsNullOrEmpty(remainingText))
        {
            // The paragraph only contained the placeholder — replace it entirely
            var parent = paragraph.Parent!;

            // Insert all converted elements before the placeholder paragraph
            OpenXmlElement insertAfter = paragraph;
            foreach (var element in elements)
            {
                parent.InsertAfter(element, insertAfter);
                insertAfter = element;
            }

            // Remove the now-empty placeholder paragraph
            paragraph.Remove();
        }
        else
        {
            // Paragraph had other content — insert elements after the paragraph
            var parent = paragraph.Parent!;
            OpenXmlElement insertAfter = paragraph;
            foreach (var element in elements)
            {
                parent.InsertAfter(element, insertAfter);
                insertAfter = element;
            }
        }
    }
}
