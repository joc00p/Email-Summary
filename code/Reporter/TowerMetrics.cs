namespace Reporter;

/// <summary>
/// Server / database counts pulled from the weekly tower emails and rendered into the
/// "Managed Services Tasks" table (lower-right) of the export template.
/// A null value means the number wasn't found in this week's emails — in that case the
/// template's existing number is left untouched rather than blanked or guessed.
/// </summary>
public class TowerMetrics
{
    // SAP — instance and server counts for RISE and XETA
    public int? SapInstances { get; set; }
    public int? SapRiseServers { get; set; }
    public int? SapRiseLiveApps { get; set; }
    public int? SapXetaServers { get; set; }
    public int? SapXetaLiveApps { get; set; }

    // DBA SQL — number of databases
    public int? SqlDatabases { get; set; }

    // CLOUD — number of servers across Sandbox / DEV / PROD
    public int? CloudServers { get; set; }

    public bool AnyFound =>
        SapInstances.HasValue || SapRiseServers.HasValue || SapRiseLiveApps.HasValue ||
        SapXetaServers.HasValue || SapXetaLiveApps.HasValue ||
        SqlDatabases.HasValue || CloudServers.HasValue;
}
