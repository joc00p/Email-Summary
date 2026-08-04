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

    // Default export/report folder when the user hasn't chosen one — the folder the exe launches from.
    public static readonly string DefaultExportDir = AppContext.BaseDirectory;

    public string TemplatePath { get; set; } = DefaultTemplatePath;

    // User-chosen default folder for exports/reports (Settings). Defaults to the exe folder (see Load()).
    public string ExportDir { get; set; } = DefaultExportDir;

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

        // Default the export folder to the exe's directory when unset.
        if (string.IsNullOrWhiteSpace(settings.ExportDir))
            settings.ExportDir = DefaultExportDir;

        return settings;
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
        File.WriteAllText(SettingsPath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
    }
}
