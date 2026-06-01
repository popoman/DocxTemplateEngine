using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplateEngine.Engine;

/// <summary>
/// Represents a found placeholder in the document with its location context.
/// </summary>
public class PlaceholderMatch
{
    public string Name { get; set; } = string.Empty;
    public Paragraph Paragraph { get; set; } = null!;
    public List<Run> Runs { get; set; } = new();
    public int StartRunIndex { get; set; }
    public int StartCharIndex { get; set; }
    public int EndRunIndex { get; set; }
    public int EndCharIndex { get; set; }
}

/// <summary>
/// Finds {{...}} placeholders in DOCX documents, handling the split-run problem
/// where Word may split a placeholder across multiple XML runs.
/// </summary>
public static class PlaceholderFinder
{
    private const string OpenDelimiter = "{{";
    private const string CloseDelimiter = "}}";

    public static List<PlaceholderMatch> FindAll(OpenXmlCompositeElement container)
    {
        var results = new List<PlaceholderMatch>();
        var paragraphs = container.Descendants<Paragraph>().ToList();

        foreach (var paragraph in paragraphs)
        {
            results.AddRange(FindInParagraph(paragraph));
        }

        return results;
    }

    public static List<PlaceholderMatch> FindInParagraph(Paragraph paragraph)
    {
        var results = new List<PlaceholderMatch>();
        var runs = paragraph.Elements<Run>().ToList();
        if (runs.Count == 0) return results;

        // Concatenate all run text to find placeholders across run boundaries
        var textParts = new List<(Run Run, int RunIndex, string Text)>();
        for (int i = 0; i < runs.Count; i++)
        {
            var text = runs[i].InnerText;
            if (!string.IsNullOrEmpty(text))
                textParts.Add((runs[i], i, text));
        }

        if (textParts.Count == 0) return results;

        var fullText = string.Concat(textParts.Select(p => p.Text));
        var offset = 0;

        while (true)
        {
            var startIdx = fullText.IndexOf(OpenDelimiter, offset, StringComparison.Ordinal);
            if (startIdx < 0) break;

            var endIdx = fullText.IndexOf(CloseDelimiter, startIdx + OpenDelimiter.Length, StringComparison.Ordinal);
            if (endIdx < 0) break;

            var name = fullText.Substring(
                startIdx + OpenDelimiter.Length,
                endIdx - startIdx - OpenDelimiter.Length).Trim();

            var (startRun, startRunIdx, startCharIdx) = MapPositionToRun(textParts, startIdx);
            var (endRun, endRunIdx, endCharIdx) = MapPositionToRun(textParts, endIdx + CloseDelimiter.Length - 1);

            var matchRuns = new List<Run>();
            for (int i = startRunIdx; i <= endRunIdx; i++)
                matchRuns.Add(textParts[i].Run);

            results.Add(new PlaceholderMatch
            {
                Name = name,
                Paragraph = paragraph,
                Runs = matchRuns,
                StartRunIndex = startRunIdx,
                StartCharIndex = startCharIdx,
                EndRunIndex = endRunIdx,
                EndCharIndex = endCharIdx
            });

            offset = endIdx + CloseDelimiter.Length;
        }

        return results;
    }

    private static (Run Run, int RunIndex, int CharIndex) MapPositionToRun(
        List<(Run Run, int RunIndex, string Text)> textParts, int globalPos)
    {
        int cumulative = 0;
        foreach (var (run, runIndex, text) in textParts)
        {
            if (globalPos < cumulative + text.Length)
                return (run, runIndex, globalPos - cumulative);
            cumulative += text.Length;
        }
        var last = textParts[^1];
        return (last.Run, last.RunIndex, last.Text.Length - 1);
    }

    /// <summary>
    /// Removes the placeholder text from the runs, merging/splitting as needed.
    /// Returns the first run (cleared) where replacement content should be inserted.
    /// </summary>
    public static Run RemovePlaceholderText(PlaceholderMatch match)
    {
        var runs = match.Paragraph.Elements<Run>().ToList();
        var textParts = new List<(Run Run, int RunIndex, string Text)>();
        for (int i = 0; i < runs.Count; i++)
        {
            var text = runs[i].InnerText;
            if (!string.IsNullOrEmpty(text))
                textParts.Add((runs[i], i, text));
        }

        // Calculate global start/end positions
        int globalStart = 0;
        for (int i = 0; i < match.StartRunIndex; i++)
            globalStart += textParts[i].Text.Length;
        globalStart += match.StartCharIndex;

        int globalEnd = 0;
        for (int i = 0; i < match.EndRunIndex; i++)
            globalEnd += textParts[i].Text.Length;
        globalEnd += match.EndCharIndex;

        // Process each run involved
        for (int i = match.StartRunIndex; i <= match.EndRunIndex; i++)
        {
            var (run, _, text) = textParts[i];
            int runGlobalStart = 0;
            for (int j = 0; j < i; j++)
                runGlobalStart += textParts[j].Text.Length;

            int localStart = Math.Max(0, globalStart - runGlobalStart);
            int localEnd = Math.Min(text.Length - 1, globalEnd - runGlobalStart);

            var newText = text.Remove(localStart, localEnd - localStart + 1);

            var textElement = run.GetFirstChild<Text>();
            if (textElement != null)
            {
                textElement.Text = newText;
                if (!string.IsNullOrEmpty(newText))
                    textElement.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        // Remove completely emptied middle runs
        for (int i = match.StartRunIndex + 1; i <= match.EndRunIndex; i++)
        {
            var run = textParts[i].Run;
            if (string.IsNullOrEmpty(run.InnerText))
                run.Remove();
        }

        return textParts[match.StartRunIndex].Run;
    }
}
