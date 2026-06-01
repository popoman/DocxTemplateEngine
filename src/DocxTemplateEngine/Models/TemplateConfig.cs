using System.Text.Json;
using System.Text.Json.Serialization;

namespace DocxTemplateEngine.Models;

public class TemplateConfig
{
    [JsonPropertyName("placeholders")]
    public Dictionary<string, PlaceholderEntry> Placeholders { get; set; } = new();

    public static TemplateConfig Load(string configPath)
    {
        if (!System.IO.File.Exists(configPath))
            throw new FileNotFoundException($"Configuration file not found: {configPath}");

        var json = System.IO.File.ReadAllText(configPath);
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
        };

        var config = JsonSerializer.Deserialize<TemplateConfig>(json, options)
            ?? throw new InvalidOperationException("Failed to deserialize configuration file.");

        var configDir = Path.GetDirectoryName(Path.GetFullPath(configPath)) ?? ".";
        foreach (var entry in config.Placeholders.Values)
        {
            if (!Path.IsPathRooted(entry.Source))
                entry.Source = Path.GetFullPath(Path.Combine(configDir, entry.Source));
        }

        config.Validate();
        return config;
    }

    public void Validate()
    {
        foreach (var (name, entry) in Placeholders)
        {
            if (string.IsNullOrWhiteSpace(entry.Source))
                throw new InvalidOperationException($"Placeholder '{name}' has an empty source path.");

            if (!System.IO.File.Exists(entry.Source))
                throw new FileNotFoundException($"Source file for placeholder '{name}' not found: {entry.Source}");

            if (entry.Type == PlaceholderType.Image)
            {
                var ext = Path.GetExtension(entry.Source).ToLowerInvariant();
                if (ext is not (".png" or ".jpg" or ".jpeg"))
                    throw new InvalidOperationException(
                        $"Placeholder '{name}' has unsupported image format '{ext}'. Supported: .png, .jpg, .jpeg");
            }
        }
    }
}
