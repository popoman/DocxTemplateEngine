using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using MarkdigTable = Markdig.Extensions.Tables.Table;
using MarkdigTableRow = Markdig.Extensions.Tables.TableRow;
using MarkdigTableCell = Markdig.Extensions.Tables.TableCell;

namespace DocxTemplateEngine.Converters;

/// <summary>
/// Converts a Markdig pipe table into an OpenXml Table element.
/// </summary>
public class MarkdownTableConverter
{
    private readonly WordprocessingDocument _document;

    public MarkdownTableConverter(WordprocessingDocument document)
    {
        _document = document;
    }

    public Table? Convert(string markdown)
    {
        var pipeline = new MarkdownPipelineBuilder()
            .UsePipeTables()
            .Build();

        var doc = Markdown.Parse(markdown, pipeline);

        // Find the first table in the parsed document
        MarkdigTable? markdigTable = null;
        foreach (var block in doc)
        {
            if (block is MarkdigTable t)
            {
                markdigTable = t;
                break;
            }
        }
        if (markdigTable == null)
            return null;

        return ConvertTable(markdigTable);
    }

    /// <summary>
    /// Converts an already-parsed Markdig table block into an OpenXml Table.
    /// </summary>
    public Table ConvertTable(MarkdigTable markdigTable)
    {
        var table = new Table();

        // Table properties: borders, width
        var tblProps = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "auto" }
            ),
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
        );
        table.Append(tblProps);

        bool isFirstRow = true;
        foreach (var row in markdigTable)
        {
            if (row is MarkdigTableRow markdigRow)
            {
                var tableRow = ConvertRow(markdigRow, isFirstRow);
                table.Append(tableRow);
                isFirstRow = false;
            }
        }

        return table;
    }

    private TableRow ConvertRow(MarkdigTableRow markdigRow, bool isHeader)
    {
        var row = new TableRow();

        if (isHeader)
        {
            var trProps = new TableRowProperties(
                new TableHeader()
            );
            row.Append(trProps);
        }

        foreach (var cell in markdigRow)
        {
            if (cell is MarkdigTableCell markdigCell)
            {
                row.Append(ConvertCell(markdigCell, isHeader));
            }
        }

        return row;
    }

    private TableCell ConvertCell(MarkdigTableCell markdigCell, bool isHeader)
    {
        var cell = new TableCell();

        // Shading for header row
        if (isHeader)
        {
            var tcProps = new TableCellProperties(
                new Shading
                {
                    Val = ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = "D9E2F3"
                }
            );
            cell.Append(tcProps);
        }

        var paragraph = new Paragraph();
        var inlineConverter = new MarkdownToOpenXmlConverter(_document);

        // Extract inline content from the cell
        foreach (var block in markdigCell)
        {
            if (block is Markdig.Syntax.ParagraphBlock pb && pb.Inline != null)
            {
                var cellContent = ExtractInlineText(pb.Inline);
                var run = new Run();

                if (isHeader)
                {
                    run.Append(new RunProperties(new Bold()));
                }

                run.Append(new Text(cellContent) { Space = SpaceProcessingModeValues.Preserve });
                paragraph.Append(run);
            }
        }

        cell.Append(paragraph);
        return cell;
    }

    /// <summary>
    /// Extracts only the data rows (skipping header) from a markdown pipe table.
    /// Each row is a list of cell text values.
    /// </summary>
    public List<List<string>> ExtractDataRows(string markdown)
    {
        var pipeline = new MarkdownPipelineBuilder()
            .UsePipeTables()
            .Build();

        var doc = Markdown.Parse(markdown, pipeline);

        MarkdigTable? markdigTable = null;
        foreach (var block in doc)
        {
            if (block is MarkdigTable t)
            {
                markdigTable = t;
                break;
            }
        }

        if (markdigTable == null)
            return [];

        var rows = new List<List<string>>();
        bool isFirstRow = true;

        foreach (var row in markdigTable)
        {
            if (row is MarkdigTableRow markdigRow)
            {
                if (isFirstRow)
                {
                    // Skip the header row — the existing DOCX table already has one
                    isFirstRow = false;
                    continue;
                }

                var cells = new List<string>();
                foreach (var cell in markdigRow)
                {
                    if (cell is MarkdigTableCell markdigCell)
                    {
                        var text = string.Empty;
                        foreach (var block in markdigCell)
                        {
                            if (block is Markdig.Syntax.ParagraphBlock pb && pb.Inline != null)
                                text = ExtractInlineText(pb.Inline);
                        }
                        cells.Add(text);
                    }
                }
                rows.Add(cells);
            }
        }

        return rows;
    }

    private static string ExtractInlineText(Markdig.Syntax.Inlines.ContainerInline container)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var inline in container)
        {
            if (inline is Markdig.Syntax.Inlines.LiteralInline literal)
                sb.Append(literal.Content);
            else if (inline is Markdig.Syntax.Inlines.ContainerInline nested)
                sb.Append(ExtractInlineText(nested));
            else if (inline is Markdig.Syntax.Inlines.CodeInline code)
                sb.Append(code.Content);
        }
        return sb.ToString();
    }
}
