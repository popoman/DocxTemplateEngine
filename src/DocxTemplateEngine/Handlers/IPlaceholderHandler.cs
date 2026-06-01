using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplateEngine.Handlers;

public interface IPlaceholderHandler
{
    /// <summary>
    /// Replaces a placeholder in the document with the appropriate content.
    /// </summary>
    /// <param name="match">The placeholder match containing location info.</param>
    /// <param name="source">The resolved file path to the source data.</param>
    /// <param name="document">The WordprocessingDocument for adding parts.</param>
    /// <param name="options">Additional options from the placeholder config.</param>
    void Replace(
        Engine.PlaceholderMatch match,
        string source,
        DocumentFormat.OpenXml.Packaging.WordprocessingDocument document,
        Models.PlaceholderEntry entry);
}
