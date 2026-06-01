using DocxTemplateEngine.Engine;
using DocxTemplateEngine.Models;

return Run(args);

static int Run(string[] args)
{
    if (args.Length == 0 || args.Contains("--help") || args.Contains("-h"))
    {
        PrintHelp();
        return 0;
    }

    var template = GetArg(args, "--template", "-t");
    var config = GetArg(args, "--config", "-c");
    var output = GetArg(args, "--output", "-o");
    var verbose = HasFlag(args, "--verbose", "-v");
    var dryRun = HasFlag(args, "--dry-run");

    if (string.IsNullOrEmpty(template))
    {
        Console.Error.WriteLine("Error: --template (-t) is required.");
        return 1;
    }
    if (string.IsNullOrEmpty(config))
    {
        Console.Error.WriteLine("Error: --config (-c) is required.");
        return 1;
    }
    if (!dryRun && string.IsNullOrEmpty(output))
    {
        Console.Error.WriteLine("Error: --output (-o) is required (unless --dry-run is used).");
        return 1;
    }

    try
    {
        if (!File.Exists(template))
        {
            Console.Error.WriteLine($"Error: Template file not found: {template}");
            return 1;
        }
        if (!File.Exists(config))
        {
            Console.Error.WriteLine($"Error: Config file not found: {config}");
            return 1;
        }

        if (output != null)
        {
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(output));
            if (outputDir != null && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);
        }

        if (verbose)
        {
            Console.WriteLine("DOCX Template Engine");
            Console.WriteLine($"  Template: {Path.GetFullPath(template)}");
            Console.WriteLine($"  Config:   {Path.GetFullPath(config)}");
            if (output != null)
                Console.WriteLine($"  Output:   {Path.GetFullPath(output)}");
            Console.WriteLine();
        }

        var templateConfig = TemplateConfig.Load(config);

        if (verbose)
        {
            Console.WriteLine($"Loaded {templateConfig.Placeholders.Count} placeholder(s) from config.");
            foreach (var (name, entry) in templateConfig.Placeholders)
                Console.WriteLine($"  {name}: {entry.Type} -> {entry.Source}");
            Console.WriteLine();
        }

        var processor = new TemplateProcessor(verbose);

        if (dryRun)
        {
            Console.WriteLine("Running in dry-run mode (no output will be generated).");
            Console.WriteLine();
            var result = processor.DryRun(template, templateConfig);
            result.Print();
            return result.IsValid ? 0 : 1;
        }
        else
        {
            processor.Process(template, templateConfig, output!);
            return 0;
        }
    }
    catch (FileNotFoundException ex)
    {
        Console.Error.WriteLine($"Error: {ex.Message}");
        return 1;
    }
    catch (InvalidOperationException ex)
    {
        Console.Error.WriteLine($"Error: {ex.Message}");
        return 2;
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        if (verbose)
            Console.Error.WriteLine(ex.StackTrace);
        return 2;
    }
}

static string? GetArg(string[] args, string longName, string? shortName = null)
{
    for (int i = 0; i < args.Length - 1; i++)
    {
        if (args[i] == longName || (shortName != null && args[i] == shortName))
            return args[i + 1];
    }
    return null;
}

static bool HasFlag(string[] args, string longName, string? shortName = null)
{
    return args.Contains(longName) || (shortName != null && args.Contains(shortName));
}

static void PrintHelp()
{
    Console.WriteLine("DOCX Template Engine - Fill DOCX templates with content from markdown, images, files, and tables.");
    Console.WriteLine();
    Console.WriteLine("Usage:");
    Console.WriteLine("  docx-template-engine --template <path.docx> --config <config.json> --output <output.docx>");
    Console.WriteLine();
    Console.WriteLine("Options:");
    Console.WriteLine("  -t, --template  (required)  Path to the DOCX template file");
    Console.WriteLine("  -c, --config    (required)  Path to the JSON configuration file");
    Console.WriteLine("  -o, --output    (required)  Path for the generated output DOCX");
    Console.WriteLine("  -v, --verbose               Enable verbose logging");
    Console.WriteLine("      --dry-run               Validate config and template without generating output");
    Console.WriteLine("  -h, --help                  Show this help message");
}
