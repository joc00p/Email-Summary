using System;
using System.IO;
using System.Text.Json;

namespace Reporter;

public class AppSettings
{
    private static readonly string SettingsPath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                     "RTXReporter", "settings.json");

    public static readonly string DefaultTemplatePath =
        Path.Combine(AppContext.BaseDirectory, "NEWTEMPLATE.pptx");

    public string TemplatePath { get; set; } = DefaultTemplatePath;

    // User-chosen default folder for exports/reports (Settings). Blank = ask each time / use last-used folder.
    public string ExportDir { get; set; } = "";

    // Folder of the most recent PPTX export — used to default the Save dialog and the "Open File Location" menu.
    public string LastExportDir { get; set; } = "";

    // The folder exports should default to: the configured ExportDir if set/valid, else the last-used folder.
    public string EffectiveExportDir =>
        !string.IsNullOrWhiteSpace(ExportDir) && Directory.Exists(ExportDir) ? ExportDir : LastExportDir;

    public static AppSettings Load()
    {
        var settings = new AppSettings();
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
        }
        catch { }

        // Fall back to the bundled template if the saved path is blank or no longer exists.
        if (string.IsNullOrWhiteSpace(settings.TemplatePath) || !File.Exists(settings.TemplatePath))
            settings.TemplatePath = DefaultTemplatePath;

        return settings;
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
        File.WriteAllText(SettingsPath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
    }
}
