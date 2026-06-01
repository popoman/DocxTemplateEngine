using System.Text.Json;
using System.Text.Json.Serialization;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace DocxTemplateEngine.Models;

public class TemplateConfig
{
    [JsonPropertyName("placeholders")]
    public Dictionary<string, PlaceholderEntry> Placeholders { get; set; } = new();

    public static TemplateConfig Load(string configPath)
    {
        if (!System.IO.File.Exists(configPath))
            throw new FileNotFoundException($"Configuration file not found: {configPath}");

        var ext = Path.GetExtension(configPath).ToLowerInvariant();
        var content = System.IO.File.ReadAllText(configPath);

        var config = ext switch
        {
            ".yaml" or ".yml" => DeserializeYaml(content),
            ".json" => DeserializeJson(content),
            _ => throw new InvalidOperationException(
                $"Unsupported config file format '{ext}'. Supported: .json, .yaml, .yml")
        };

        var configDir = Path.GetDirectoryName(Path.GetFullPath(configPath)) ?? ".";
        foreach (var entry in config.Placeholders.Values)
        {
            if (!Path.IsPathRooted(entry.Source))
                entry.Source = Path.GetFullPath(Path.Combine(configDir, entry.Source));
        }

        config.Validate();
        return config;
    }

    private static TemplateConfig DeserializeJson(string json)
    {
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
        };

        return JsonSerializer.Deserialize<TemplateConfig>(json, options)
            ?? throw new InvalidOperationException("Failed to deserialize JSON configuration file.");
    }

    private static TemplateConfig DeserializeYaml(string yaml)
    {
        var deserializer = new DeserializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .Build();

        return deserializer.Deserialize<TemplateConfig>(yaml)
            ?? throw new InvalidOperationException("Failed to deserialize YAML configuration file.");
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
