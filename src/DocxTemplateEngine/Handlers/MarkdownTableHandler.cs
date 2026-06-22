using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Converters;
using DocxTemplateEngine.Engine;
using DocxTemplateEngine.Models;

namespace DocxTemplateEngine.Handlers;

public class MarkdownTableHandler : IPlaceholderHandler
{
    public void Replace(PlaceholderMatch match, string source, WordprocessingDocument document, PlaceholderEntry entry)
    {
        var markdown = File.ReadAllText(source);

        // Auto-detect: is the placeholder inside an existing table?
        var parentCell = FindAncestor<TableCell>(match.Paragraph);
        if (parentCell != null)
        {
            PopulateExistingTable(match, markdown, document, parentCell);
        }
        else
        {
            CreateNewTable(match, markdown, document);
        }
    }

    private void CreateNewTable(PlaceholderMatch match, string markdown, WordprocessingDocument document)
    {
        var converter = new MarkdownTableConverter();
        var table = converter.Convert(markdown);

        if (table == null)
        {
            Console.WriteLine($"Warning: No table found in markdown source.");
            return;
        }

        PlaceholderFinder.RemovePlaceholderText(match);

        var paragraph = match.Paragraph;
        var parent = paragraph.Parent!;

        var remainingText = paragraph.InnerText.Trim();
        if (string.IsNullOrEmpty(remainingText))
        {
            parent.InsertAfter(table, paragraph);
            paragraph.Remove();
        }
        else
        {
            parent.InsertAfter(table, paragraph);
        }
    }

    private void PopulateExistingTable(PlaceholderMatch match, string markdown,
        WordprocessingDocument document, TableCell placeholderCell)
    {
        var converter = new MarkdownTableConverter();
        var dataRows = converter.ExtractDataRows(markdown);

        if (dataRows.Count == 0)
        {
            Console.WriteLine("Warning: No data rows found in markdown table.");
            return;
        }

        var placeholderRow = FindAncestor<TableRow>(placeholderCell);
        var existingTable = FindAncestor<Table>(placeholderCell);

        if (existingTable == null || placeholderRow == null)
        {
            Console.WriteLine("Warning: Could not locate parent table for existing-table population. Falling back to new table.");
            CreateNewTable(match, markdown, document);
            return;
        }

        // Clone cell properties from the placeholder row to preserve column widths/borders
        var templateCellProps = placeholderRow.Elements<TableCell>()
            .Select(c => c.GetFirstChild<TableCellProperties>()?.CloneNode(true) as TableCellProperties)
            .ToList();

        // Insert data rows after the placeholder row
        OpenXmlElement insertAfter = placeholderRow;
        foreach (var rowData in dataRows)
        {
            var newRow = new TableRow();

            for (int colIdx = 0; colIdx < rowData.Count; colIdx++)
            {
                var cell = new TableCell();

                // Clone cell properties from the template row if available
                if (colIdx < templateCellProps.Count && templateCellProps[colIdx] != null)
                {
                    cell.Append(templateCellProps[colIdx]!.CloneNode(true));
                }

                var para = new Paragraph(
                    new Run(new Text(rowData[colIdx]) { Space = SpaceProcessingModeValues.Preserve })
                );
                cell.Append(para);
                newRow.Append(cell);
            }

            existingTable.InsertAfter(newRow, insertAfter);
            insertAfter = newRow;
        }

        // Remove the placeholder row
        placeholderRow.Remove();
    }

    private static T? FindAncestor<T>(OpenXmlElement? element) where T : OpenXmlElement
    {
        var current = element?.Parent;
        while (current != null)
        {
            if (current is T found)
                return found;
            current = current.Parent;
        }
        return null;
    }
}
