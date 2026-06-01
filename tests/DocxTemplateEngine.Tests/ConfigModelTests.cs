using System.Text.Json;
using DocxTemplateEngine.Models;
using FluentAssertions;

namespace DocxTemplateEngine.Tests;

public class ConfigModelTests : IDisposable
{
    private readonly string _tempDir;

    public ConfigModelTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"DocxTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public void Load_ValidConfig_DeserializesCorrectly()
    {
        // Create a source file referenced by the config
        var mdPath = Path.Combine(_tempDir, "intro.md");
        File.WriteAllText(mdPath, "# Hello");

        var configJson = $$"""
        {
          "placeholders": {
            "Introduction": {
              "type": "markdown",
              "source": "intro.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var config = TemplateConfig.Load(configPath);

        config.Placeholders.Should().HaveCount(1);
        config.Placeholders["Introduction"].Type.Should().Be(PlaceholderType.Markdown);
        config.Placeholders["Introduction"].Source.Should().EndWith("intro.md");
    }

    [Fact]
    public void Load_MissingSourceFile_ThrowsFileNotFound()
    {
        var configJson = """
        {
          "placeholders": {
            "Test": {
              "type": "markdown",
              "source": "nonexistent.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var act = () => TemplateConfig.Load(configPath);
        act.Should().Throw<FileNotFoundException>();
    }

    [Fact]
    public void Load_InvalidImageFormat_ThrowsInvalidOperation()
    {
        var bmpPath = Path.Combine(_tempDir, "image.bmp");
        File.WriteAllText(bmpPath, "fake bmp");

        var configJson = """
        {
          "placeholders": {
            "Logo": {
              "type": "image",
              "source": "image.bmp"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var act = () => TemplateConfig.Load(configPath);
        act.Should().Throw<InvalidOperationException>()
            .WithMessage("*unsupported image format*");
    }

    [Fact]
    public void Load_ResolvesRelativePaths()
    {
        var subDir = Path.Combine(_tempDir, "content");
        Directory.CreateDirectory(subDir);
        File.WriteAllText(Path.Combine(subDir, "data.md"), "# Data");

        var configJson = """
        {
          "placeholders": {
            "Data": {
              "type": "markdown",
              "source": "content/data.md"
            }
          }
        }
        """;

        var configPath = Path.Combine(_tempDir, "config.json");
        File.WriteAllText(configPath, configJson);

        var config = TemplateConfig.Load(configPath);

        config.Placeholders["Data"].Source.Should().Be(
            Path.GetFullPath(Path.Combine(_tempDir, "content", "data.md")));
    }
}
