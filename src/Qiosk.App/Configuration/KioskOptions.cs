namespace Qiosk.App.Configuration;

public sealed class KioskOptions
{
    public const string SectionName = "Kiosk";

    public string ExcelPath { get; set; } = "Data\\attendees.xlsx";
    public string TemplatePath { get; set; } = "Templates\\badge-template.docx";
    public int CooldownSeconds { get; set; } = 3;
    public int PrintingMessageSeconds { get; set; } = 10;
    public string BeepSound { get; set; } = "Asterisk";
}
