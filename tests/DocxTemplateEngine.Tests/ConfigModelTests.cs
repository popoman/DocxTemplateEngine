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

    [Fact]
    public void Load_YamlConfig_DeserializesCorrectly()
    {
        var mdPath = Path.Combine(_tempDir, "intro.md");
        File.WriteAllText(mdPath, "# Hello");

        var yaml = """
            placeholders:
              Introduction:
                type: Markdown
                source: intro.md
            """;

        var configPath = Path.Combine(_tempDir, "config.yaml");
        File.WriteAllText(configPath, yaml);

        var config = TemplateConfig.Load(configPath);

        config.Placeholders.Should().HaveCount(1);
        config.Placeholders["Introduction"].Type.Should().Be(PlaceholderType.Markdown);
        config.Placeholders["Introduction"].Source.Should().EndWith("intro.md");
    }

    [Fact]
    public void Load_YmlExtension_AlsoWorks()
    {
        var mdPath = Path.Combine(_tempDir, "data.md");
        File.WriteAllText(mdPath, "# Data");

        var yaml = """
            placeholders:
              Data:
                type: Markdown
                source: data.md
            """;

        var configPath = Path.Combine(_tempDir, "config.yml");
        File.WriteAllText(configPath, yaml);

        var config = TemplateConfig.Load(configPath);
        config.Placeholders.Should().HaveCount(1);
    }

    [Fact]
    public void Load_YamlWithAllTypes_DeserializesCorrectly()
    {
        var mdPath = Path.Combine(_tempDir, "intro.md");
        File.WriteAllText(mdPath, "# Hello");
        var tablePath = Path.Combine(_tempDir, "table.md");
        File.WriteAllText(tablePath, "| A | B |\n|---|---|\n| 1 | 2 |");

        // Minimal 1x1 PNG
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(Path.Combine(_tempDir, "logo.png"), pngBytes);
        File.WriteAllText(Path.Combine(_tempDir, "report.txt"), "report");

        var yaml = """
            placeholders:
              Content:
                type: Markdown
                source: intro.md
              Logo:
                type: Image
                source: logo.png
                widthCm: 5.0
                heightCm: 3.0
              Attachment:
                type: File
                source: report.txt
                displayName: Monthly Report
              Data:
                type: MarkdownTable
                source: table.md
            """;

        var configPath = Path.Combine(_tempDir, "config.yaml");
        File.WriteAllText(configPath, yaml);

        var config = TemplateConfig.Load(configPath);

        config.Placeholders.Should().HaveCount(4);
        config.Placeholders["Content"].Type.Should().Be(PlaceholderType.Markdown);
        config.Placeholders["Logo"].Type.Should().Be(PlaceholderType.Image);
        config.Placeholders["Logo"].WidthCm.Should().Be(5.0);
        config.Placeholders["Logo"].HeightCm.Should().Be(3.0);
        config.Placeholders["Attachment"].Type.Should().Be(PlaceholderType.File);
        config.Placeholders["Attachment"].DisplayName.Should().Be("Monthly Report");
        config.Placeholders["Data"].Type.Should().Be(PlaceholderType.MarkdownTable);
    }

    [Fact]
    public void Load_UnsupportedExtension_Throws()
    {
        var configPath = Path.Combine(_tempDir, "config.xml");
        File.WriteAllText(configPath, "<xml/>");

        var act = () => TemplateConfig.Load(configPath);
        act.Should().Throw<InvalidOperationException>()
            .WithMessage("*Unsupported config file format*");
    }

    [Fact]
    public void Load_YamlResolvesRelativePaths()
    {
        var subDir = Path.Combine(_tempDir, "content");
        Directory.CreateDirectory(subDir);
        File.WriteAllText(Path.Combine(subDir, "data.md"), "# Data");

        var yaml = """
            placeholders:
              Data:
                type: Markdown
                source: content/data.md
            """;

        var configPath = Path.Combine(_tempDir, "config.yaml");
        File.WriteAllText(configPath, yaml);

        var config = TemplateConfig.Load(configPath);
        config.Placeholders["Data"].Source.Should().Be(
            Path.GetFullPath(Path.Combine(_tempDir, "content", "data.md")));
    }
}
