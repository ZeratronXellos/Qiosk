namespace Qiosk.App.Configuration;

public sealed class KioskSettings
{
    public string DefaultPrinter { get; set; } = string.Empty;
    public string DefaultCameraMoniker { get; set; } = string.Empty;
    public bool LockMode { get; set; } = true;
    public string ExcelPathOverride { get; set; } = string.Empty;
    public string TemplatePathOverride { get; set; } = string.Empty;
    public bool ReloadOnStart { get; set; } = true;
}
