using System.Text.Json.Serialization;

namespace DocxTemplateEngine.Models;

public class PlaceholderEntry
{
    [JsonPropertyName("type")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public PlaceholderType Type { get; set; }

    [JsonPropertyName("source")]
    public string Source { get; set; } = string.Empty;

    [JsonPropertyName("widthCm")]
    public double? WidthCm { get; set; }

    [JsonPropertyName("heightCm")]
    public double? HeightCm { get; set; }

    [JsonPropertyName("displayName")]
    public string? DisplayName { get; set; }
}
