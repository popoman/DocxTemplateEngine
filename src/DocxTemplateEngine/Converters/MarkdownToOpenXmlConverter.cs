using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace DocxTemplateEngine.Converters;

/// <summary>
/// Converts Markdig AST elements into OpenXml elements for DOCX documents.
/// Supports: headings, bold, italic, inline code, code blocks, lists, links, paragraphs.
/// </summary>
public class MarkdownToOpenXmlConverter
{
    private readonly WordprocessingDocument _document;
    private int _numberingId;

    public MarkdownToOpenXmlConverter(WordprocessingDocument document)
    {
        _document = document;
    }

    public List<OpenXmlElement> Convert(string markdown)
    {
        var pipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .Build();

        var doc = Markdown.Parse(markdown, pipeline);
        var elements = new List<OpenXmlElement>();

        foreach (var block in doc)
        {
            elements.AddRange(ConvertBlock(block));
        }

        return elements;
    }

    private List<OpenXmlElement> ConvertBlock(Block block)
    {
        return block switch
        {
            HeadingBlock heading => ConvertHeading(heading),
            ParagraphBlock paragraph => ConvertParagraph(paragraph),
            FencedCodeBlock fencedCode => ConvertCodeBlock(fencedCode),
            CodeBlock codeBlock => ConvertCodeBlock(codeBlock),
            ListBlock list => ConvertList(list),
            QuoteBlock quote => ConvertQuoteBlock(quote),
            ThematicBreakBlock => ConvertThematicBreak(),
            LinkReferenceDefinitionGroup => [],  // Internal Markdig metadata, not rendered
            LinkReferenceDefinition => [],
            _ => ConvertFallbackBlock(block)
        };
    }

    private List<OpenXmlElement> ConvertHeading(HeadingBlock heading)
    {
        var paragraph = new Paragraph();
        var pProps = new ParagraphProperties
        {
            ParagraphStyleId = new ParagraphStyleId { Val = $"Heading{heading.Level}" }
        };
        paragraph.Append(pProps);

        if (heading.Inline != null)
        {
            foreach (var inline in heading.Inline)
            {
                paragraph.Append(ConvertInline(inline));
            }
        }

        return [paragraph];
    }

    private List<OpenXmlElement> ConvertParagraph(ParagraphBlock paragraphBlock)
    {
        var paragraph = new Paragraph();

        if (paragraphBlock.Inline != null)
        {
            foreach (var inline in paragraphBlock.Inline)
            {
                paragraph.Append(ConvertInline(inline));
            }
        }

        return [paragraph];
    }

    private List<OpenXmlElement> ConvertCodeBlock(CodeBlock codeBlock)
    {
        var lines = codeBlock.Lines.ToString().TrimEnd('\n', '\r');
        var paragraph = new Paragraph();

        var pProps = new ParagraphProperties
        {
            Shading = new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = "F0F0F0"
            },
            SpacingBetweenLines = new SpacingBetweenLines { Before = "100", After = "100" }
        };
        paragraph.Append(pProps);

        foreach (var line in lines.Split('\n'))
        {
            var run = new Run();
            var runProps = new RunProperties
            {
                RunFonts = new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" },
                FontSize = new FontSize { Val = "18" }
            };
            run.Append(runProps);
            run.Append(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);

            // Add line break except for last line
            if (line != lines.Split('\n').Last())
            {
                paragraph.Append(new Run(new Break()));
            }
        }

        return [paragraph];
    }

    private List<OpenXmlElement> ConvertList(ListBlock list)
    {
        var elements = new List<OpenXmlElement>();
        var isOrdered = list.IsOrdered;
        var numId = EnsureNumberingDefinition(isOrdered);
        int itemIndex = 0;

        foreach (var item in list)
        {
            if (item is ListItemBlock listItem)
            {
                elements.AddRange(ConvertListItem(listItem, numId, 0, isOrdered, ref itemIndex));
            }
        }

        return elements;
    }

    private List<OpenXmlElement> ConvertListItem(ListItemBlock listItem, int numId, int level, bool isOrdered, ref int itemIndex)
    {
        var elements = new List<OpenXmlElement>();

        foreach (var childBlock in listItem)
        {
            if (childBlock is ParagraphBlock paragraphBlock)
            {
                var paragraph = new Paragraph();

                var numProps = new NumberingProperties(
                    new NumberingLevelReference { Val = level },
                    new NumberingId { Val = numId }
                );

                var pProps = new ParagraphProperties();
                pProps.Append(numProps);
                paragraph.Append(pProps);

                if (paragraphBlock.Inline != null)
                {
                    foreach (var inline in paragraphBlock.Inline)
                    {
                        paragraph.Append(ConvertInline(inline));
                    }
                }

                elements.Add(paragraph);
                itemIndex++;
            }
            else if (childBlock is ListBlock nestedList)
            {
                var nestedNumId = EnsureNumberingDefinition(nestedList.IsOrdered);
                int nestedIndex = 0;
                foreach (var nestedItem in nestedList)
                {
                    if (nestedItem is ListItemBlock nestedListItem)
                    {
                        elements.AddRange(ConvertListItem(nestedListItem, nestedNumId, level + 1, nestedList.IsOrdered, ref nestedIndex));
                    }
                }
            }
            else
            {
                elements.AddRange(ConvertBlock(childBlock));
            }
        }

        return elements;
    }

    private List<OpenXmlElement> ConvertQuoteBlock(QuoteBlock quote)
    {
        var elements = new List<OpenXmlElement>();

        foreach (var block in quote)
        {
            var converted = ConvertBlock(block);
            foreach (var element in converted)
            {
                if (element is Paragraph p)
                {
                    var pProps = p.GetFirstChild<ParagraphProperties>() ?? new ParagraphProperties();
                    if (p.GetFirstChild<ParagraphProperties>() == null)
                        p.PrependChild(pProps);

                    pProps.Append(new Indentation { Left = "720" });
                    pProps.Append(new ParagraphBorders(
                        new LeftBorder
                        {
                            Val = BorderValues.Single,
                            Size = 12,
                            Color = "CCCCCC",
                            Space = 4
                        }
                    ));
                }
                elements.Add(element);
            }
        }

        return elements;
    }

    private List<OpenXmlElement> ConvertThematicBreak()
    {
        var paragraph = new Paragraph();
        var pProps = new ParagraphProperties
        {
            ParagraphBorders = new ParagraphBorders(
                new BottomBorder
                {
                    Val = BorderValues.Single,
                    Size = 6,
                    Color = "auto",
                    Space = 1
                }
            )
        };
        paragraph.Append(pProps);
        return [paragraph];
    }

    private List<OpenXmlElement> ConvertFallbackBlock(Block block)
    {
        var paragraph = new Paragraph();
        var run = new Run(new Text(block.ToString() ?? string.Empty)
        {
            Space = SpaceProcessingModeValues.Preserve
        });
        paragraph.Append(run);
        return [paragraph];
    }

    private IEnumerable<OpenXmlElement> ConvertInline(Inline inline)
    {
        return inline switch
        {
            LiteralInline literal => ConvertLiteral(literal),
            EmphasisInline emphasis => ConvertEmphasis(emphasis),
            CodeInline code => ConvertInlineCode(code),
            LinkInline link => ConvertLink(link),
            LineBreakInline => [new Run(new Break())],
            HtmlInline html => ConvertHtmlInline(html),
            _ => ConvertFallbackInline(inline)
        };
    }

    private IEnumerable<OpenXmlElement> ConvertLiteral(LiteralInline literal)
    {
        var run = new Run();
        run.Append(new Text(literal.Content.ToString())
        {
            Space = SpaceProcessingModeValues.Preserve
        });
        return [run];
    }

    private IEnumerable<OpenXmlElement> ConvertEmphasis(EmphasisInline emphasis)
    {
        var runs = new List<OpenXmlElement>();

        foreach (var child in emphasis)
        {
            var childElements = ConvertInline(child);
            foreach (var element in childElements)
            {
                if (element is Run run)
                {
                    var props = run.GetFirstChild<RunProperties>() ?? new RunProperties();
                    if (run.GetFirstChild<RunProperties>() == null)
                        run.PrependChild(props);

                    if (emphasis.DelimiterChar is '*' or '_')
                    {
                        if (emphasis.DelimiterCount == 2)
                            props.Append(new Bold());
                        else
                            props.Append(new Italic());
                    }
                }
                runs.Add(element);
            }
        }

        return runs;
    }

    private IEnumerable<OpenXmlElement> ConvertInlineCode(CodeInline code)
    {
        var run = new Run();
        var props = new RunProperties
        {
            RunFonts = new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" },
            FontSize = new FontSize { Val = "18" },
            Shading = new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = "F0F0F0"
            }
        };
        run.Append(props);
        run.Append(new Text(code.Content) { Space = SpaceProcessingModeValues.Preserve });
        return [run];
    }

    private IEnumerable<OpenXmlElement> ConvertLink(LinkInline link)
    {
        var elements = new List<OpenXmlElement>();

        if (link.Url == null) return elements;

        var mainPart = _document.MainDocumentPart;
        if (mainPart == null) return elements;

        var rel = mainPart.AddHyperlinkRelationship(new Uri(link.Url, UriKind.RelativeOrAbsolute), true);

        var hyperlink = new Hyperlink { Id = rel.Id };

        foreach (var child in link)
        {
            var childElements = ConvertInline(child);
            foreach (var element in childElements)
            {
                if (element is Run run)
                {
                    var props = run.GetFirstChild<RunProperties>() ?? new RunProperties();
                    if (run.GetFirstChild<RunProperties>() == null)
                        run.PrependChild(props);

                    props.Append(new RunStyle { Val = "Hyperlink" });
                    props.Append(new Color { Val = "0563C1" });
                    props.Append(new Underline { Val = UnderlineValues.Single });
                }
                hyperlink.Append(element);
            }
        }

        elements.Add(hyperlink);
        return elements;
    }

    private IEnumerable<OpenXmlElement> ConvertHtmlInline(HtmlInline html)
    {
        if (html.Tag == "<br>" || html.Tag == "<br/>" || html.Tag == "<br />")
            return [new Run(new Break())];

        return [];
    }

    private IEnumerable<OpenXmlElement> ConvertFallbackInline(Inline inline)
    {
        var run = new Run(new Text(inline.ToString() ?? string.Empty)
        {
            Space = SpaceProcessingModeValues.Preserve
        });
        return [run];
    }

    private int EnsureNumberingDefinition(bool isOrdered)
    {
        var mainPart = _document.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart;

        if (numberingPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
        }

        var numbering = numberingPart.Numbering;
        _numberingId++;

        var abstractNum = new AbstractNum(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = isOrdered ? NumberFormatValues.Decimal : NumberFormatValues.Bullet },
                new LevelText { Val = isOrdered ? "%1." : "\u2022" },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new ParagraphProperties(
                    new Indentation { Left = "720", Hanging = "360" }
                )
            )
            { LevelIndex = 0 }
        )
        { AbstractNumberId = _numberingId };

        // Add Level 1 for nested lists
        abstractNum.Append(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = isOrdered ? NumberFormatValues.LowerLetter : NumberFormatValues.Bullet },
                new LevelText { Val = isOrdered ? "%2." : "◦" },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new ParagraphProperties(
                    new Indentation { Left = "1440", Hanging = "360" }
                )
            )
            { LevelIndex = 1 }
        );

        var numInstance = new NumberingInstance(
            new AbstractNumId { Val = _numberingId }
        )
        { NumberID = _numberingId };

        // Insert before any existing instances
        var firstInstance = numbering.Elements<NumberingInstance>().FirstOrDefault();
        if (firstInstance != null)
        {
            numbering.InsertBefore(abstractNum, firstInstance);
        }
        else
        {
            numbering.Append(abstractNum);
        }
        numbering.Append(numInstance);

        return _numberingId;
    }
}
