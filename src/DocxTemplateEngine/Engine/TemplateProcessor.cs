using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplateEngine.Handlers;
using DocxTemplateEngine.Models;

namespace DocxTemplateEngine.Engine;

public class TemplateProcessor
{
    private readonly Dictionary<PlaceholderType, IPlaceholderHandler> _handlers = new()
    {
        [PlaceholderType.Markdown] = new MarkdownHandler(),
        [PlaceholderType.Image] = new ImageHandler(),
        [PlaceholderType.File] = new FileObjectHandler(),
        [PlaceholderType.MarkdownTable] = new MarkdownTableHandler()
    };

    private readonly bool _verbose;

    public TemplateProcessor(bool verbose = false)
    {
        _verbose = verbose;
    }

    public void Process(string templatePath, TemplateConfig config, string outputPath)
    {
        Log($"Opening template: {templatePath}");

        // Copy template to output
        System.IO.File.Copy(templatePath, outputPath, overwrite: true);

        using var document = WordprocessingDocument.Open(outputPath, isEditable: true);
        var mainPart = document.MainDocumentPart
            ?? throw new InvalidOperationException("Document does not have a main part.");

        var body = mainPart.Document.Body
            ?? throw new InvalidOperationException("Document body is missing.");

        // Find all placeholders in the document body
        var matches = PlaceholderFinder.FindAll(body);
        Log($"Found {matches.Count} placeholder(s) in document body.");

        // Also search headers and footers
        foreach (var headerPart in mainPart.HeaderParts)
        {
            if (headerPart.Header != null)
            {
                var headerMatches = PlaceholderFinder.FindAll(headerPart.Header);
                matches.AddRange(headerMatches);
                Log($"Found {headerMatches.Count} placeholder(s) in a header.");
            }
        }

        foreach (var footerPart in mainPart.FooterParts)
        {
            if (footerPart.Footer != null)
            {
                var footerMatches = PlaceholderFinder.FindAll(footerPart.Footer);
                matches.AddRange(footerMatches);
                Log($"Found {footerMatches.Count} placeholder(s) in a footer.");
            }
        }

        // Process each placeholder (reverse order to avoid position shifts)
        var processedCount = 0;
        var unmatchedNames = new HashSet<string>();

        // Process in reverse order to avoid invalidating positions
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            var match = matches[i];

            if (!config.Placeholders.TryGetValue(match.Name, out var entry))
            {
                unmatchedNames.Add(match.Name);
                continue;
            }

            Log($"Processing placeholder '{{{{{match.Name}}}}}' (type: {entry.Type}) ...");

            if (!_handlers.TryGetValue(entry.Type, out var handler))
            {
                Console.WriteLine($"Warning: No handler for placeholder type '{entry.Type}'.");
                continue;
            }

            try
            {
                handler.Replace(match, entry.Source, document, entry);
                processedCount++;
                Log($"  -> Replaced successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing placeholder '{match.Name}': {ex.Message}");
                if (_verbose)
                    Console.WriteLine(ex.StackTrace);
            }
        }

        // Save
        mainPart.Document.Save();

        Log($"Processed {processedCount} placeholder(s).");

        if (unmatchedNames.Count > 0)
        {
            Console.WriteLine($"Warning: {unmatchedNames.Count} placeholder(s) not found in config: {string.Join(", ", unmatchedNames)}");
        }

        Console.WriteLine($"Output saved to: {outputPath}");
    }

    public ValidationResult DryRun(string templatePath, TemplateConfig config)
    {
        var result = new ValidationResult();

        if (!System.IO.File.Exists(templatePath))
        {
            result.Errors.Add($"Template file not found: {templatePath}");
            return result;
        }

        using var document = WordprocessingDocument.Open(templatePath, isEditable: false);
        var mainPart = document.MainDocumentPart;
        if (mainPart?.Document.Body == null)
        {
            result.Errors.Add("Document has no body content.");
            return result;
        }

        var matches = PlaceholderFinder.FindAll(mainPart.Document.Body);

        foreach (var headerPart in mainPart.HeaderParts)
        {
            if (headerPart.Header != null)
                matches.AddRange(PlaceholderFinder.FindAll(headerPart.Header));
        }

        foreach (var footerPart in mainPart.FooterParts)
        {
            if (footerPart.Footer != null)
                matches.AddRange(PlaceholderFinder.FindAll(footerPart.Footer));
        }

        result.PlaceholdersFound = matches.Select(m => m.Name).Distinct().ToList();

        foreach (var name in result.PlaceholdersFound)
        {
            if (!config.Placeholders.ContainsKey(name))
                result.Warnings.Add($"Placeholder '{name}' found in template but not in config.");
        }

        foreach (var (name, entry) in config.Placeholders)
        {
            if (!result.PlaceholdersFound.Contains(name))
                result.Warnings.Add($"Config entry '{name}' not found in template.");

            if (!System.IO.File.Exists(entry.Source))
                result.Errors.Add($"Source file for '{name}' not found: {entry.Source}");
        }

        return result;
    }

    private void Log(string message)
    {
        if (_verbose)
            Console.WriteLine($"[INFO] {message}");
    }
}

public class ValidationResult
{
    public List<string> PlaceholdersFound { get; set; } = new();
    public List<string> Errors { get; set; } = new();
    public List<string> Warnings { get; set; } = new();

    public bool IsValid => Errors.Count == 0;

    public void Print()
    {
        Console.WriteLine($"Placeholders found: {PlaceholdersFound.Count}");
        foreach (var p in PlaceholdersFound)
            Console.WriteLine($"  - {p}");

        if (Warnings.Count > 0)
        {
            Console.WriteLine($"\nWarnings ({Warnings.Count}):");
            foreach (var w in Warnings)
                Console.WriteLine($"  ⚠ {w}");
        }

        if (Errors.Count > 0)
        {
            Console.WriteLine($"\nErrors ({Errors.Count}):");
            foreach (var e in Errors)
                Console.WriteLine($"  ✗ {e}");
        }
        else
        {
            Console.WriteLine("\n✓ Validation passed.");
        }
    }
}
