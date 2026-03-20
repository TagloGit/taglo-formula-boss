using System.Text.Json;
using System.Text.Json.Serialization;

namespace FormulaBoss.UI;

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum AnimationStyle
{
    Chomp,
    Roar,
    Shuffle,
    None
}

public class EditorSettings
{
    private static readonly JsonSerializerOptions SerializerOptions = new() { WriteIndented = true };

    public double Width { get; set; } = 500;
    public double Height { get; set; } = 300;
    public int IndentSize { get; set; } = 2;
    public double FontSize { get; set; } = 13;
    public AnimationStyle AnimationStyle { get; set; } = AnimationStyle.Chomp;
    public bool WordWrap { get; set; } = true;
    public bool AutoFormatLet { get; set; } = true;
    public int NestedLetDepth { get; set; } = 1;
    public int MaxLineLength { get; set; }

    private static string SettingsDirectory =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "FormulaBoss");

    private static string SettingsPath =>
        Path.Combine(SettingsDirectory, "editor-settings.json");

    public static EditorSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                return JsonSerializer.Deserialize<EditorSettings>(json) ?? new EditorSettings();
            }
        }
        catch
        {
            // If loading fails, return defaults
        }

        return new EditorSettings();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(SettingsDirectory);
            var json = JsonSerializer.Serialize(this, SerializerOptions);
            File.WriteAllText(SettingsPath, json);
        }
        catch
        {
            // Silently fail if we can't save settings
        }
    }
}
