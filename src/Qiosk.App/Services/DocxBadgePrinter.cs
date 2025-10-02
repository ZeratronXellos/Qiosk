using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Qiosk.App.Models;
using Qiosk.App.Services.Contracts;
using Word = Microsoft.Office.Interop.Word;

namespace Qiosk.App.Services;

public sealed class DocxBadgePrinter : IBadgePrinter
{
    private readonly ILogger<DocxBadgePrinter> _logger;

    public DocxBadgePrinter(ILogger<DocxBadgePrinter> logger)
    {
        _logger = logger;
    }

    public async Task PrintAsync(Attendee attendee, string printerName, string templatePath, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(attendee);
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);

        var resolvedTemplate = Path.GetFullPath(templatePath);
        if (!File.Exists(resolvedTemplate))
        {
            throw new FileNotFoundException($"Nu am gasit sablonul DOCX la calea '{resolvedTemplate}'.", resolvedTemplate);
        }

        var tempDirectory = Path.Combine(Path.GetTempPath(), "Qiosk");
        Directory.CreateDirectory(tempDirectory);
        var tempDocPath = Path.Combine(tempDirectory, $"badge-{Guid.NewGuid():N}.docx");

        File.Copy(resolvedTemplate, tempDocPath, overwrite: true);

        try
        {
            EnsureWordInstalled();
            await PrintWithWordAsync(tempDocPath, printerName, attendee, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            TryDeleteFile(tempDocPath);
        }
    }

    private static Task PrintWithWordAsync(string docPath, string printerName, Attendee attendee, CancellationToken cancellationToken)
    {
        return RunOnStaThreadAsync(() =>
        {
            Word.Application? application = null;
            Word.Document? document = null;
            Word.Window? activeWindow = null;

            try
            {
                application = new Word.Application
                {
                    Visible = true,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
                    ScreenUpdating = false
                };

                if (!string.IsNullOrWhiteSpace(printerName))
                {
                    application.ActivePrinter = printerName;
                }

                document = application.Documents.Open(
                    docPath,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: true);

                document.Activate();

                activeWindow = document.ActiveWindow;
                if (activeWindow is not null)
                {
                    activeWindow.WindowState = Word.WdWindowState.wdWindowStateMinimize;
                }

                ApplyReplacements(document, attendee);

                document.Save();
                document.PrintOut(Background: false);
            }
            finally
            {
                if (activeWindow is not null)
                {
                    Marshal.FinalReleaseComObject(activeWindow);
                }

                if (document is not null)
                {
                    document.Close(false);
                    Marshal.FinalReleaseComObject(document);
                }

                if (application is not null)
                {
                    application.Visible = false;
                    application.Quit(false);
                    Marshal.FinalReleaseComObject(application);
                }
            }
        }, cancellationToken);
    }

    private static void ApplyReplacements(Word.Document document, Attendee attendee)
    {
        var replacements = new Dictionary<string, string?>
        {
            ["{{ID}}"] = attendee.Id,
            ["{{Nume}}"] = attendee.LastName,
            ["{{Prenume}}"] = attendee.FirstName,
            ["{{Rol}}"] = attendee.Role,
            ["{{Companie}}"] = attendee.Company,
        };

        foreach (var (placeholder, value) in replacements)
        {
            ReplaceAll(document, placeholder, value ?? string.Empty);
        }
    }

    private static void ReplaceAll(Word.Document document, string placeholder, string value)
    {
        Word.Range? range = null;
        Word.Find? find = null;

        try
        {
            range = document.Content;
            find = range.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Execute(
                FindText: placeholder,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: false,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false,
                ReplaceWith: value,
                Replace: Word.WdReplace.wdReplaceAll);
        }
        finally
        {
            if (find is not null)
            {
                Marshal.FinalReleaseComObject(find);
            }

            if (range is not null)
            {
                Marshal.FinalReleaseComObject(range);
            }
        }
    }

    private static void EnsureWordInstalled()
    {
        if (Type.GetTypeFromProgID("Word.Application") is null)
        {
            throw new InvalidOperationException("Microsoft Word nu este instalat sau nu expune COM interop. Instaleaza Word (de preferat Office 2013 sau mai nou, 64-bit) ori componentele Office Primary Interop Assemblies.");
        }
    }

    private static Task RunOnStaThreadAsync(Action action, CancellationToken cancellationToken)
    {
        if (cancellationToken.IsCancellationRequested)
        {
            return Task.FromCanceled(cancellationToken);
        }

        var completionSource = new TaskCompletionSource<object?>();

        var thread = new Thread(() =>
        {
            try
            {
                action();
                completionSource.TrySetResult(null);
            }
            catch (Exception ex)
            {
                completionSource.TrySetException(ex);
            }
        })
        {
            IsBackground = true
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        if (cancellationToken.CanBeCanceled)
        {
            cancellationToken.Register(() => completionSource.TrySetCanceled(cancellationToken));
        }

        return completionSource.Task;
    }

    private void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Nu am reusit sa sterg fisierul temporar {Path}", path);
        }
    }
}
