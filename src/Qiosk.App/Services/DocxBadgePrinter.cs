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
using Xceed.Words.NET;

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
            using var document = DocX.Load(tempDocPath);
            ApplyReplacements(document, attendee);
            document.Save();

            await PrintWithWordAsync(tempDocPath, printerName, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            TryDeleteFile(tempDocPath);
        }
    }

    private static void ApplyReplacements(DocX document, Attendee attendee)
    {
        var replacements = new Dictionary<string, string>
        {
            ["{{ID}}"] = attendee.Id,
            ["{{Nume}}"] = attendee.LastName,
            ["{{Prenume}}"] = attendee.FirstName,
            ["{{Rol}}"] = attendee.Role,
            ["{{Companie}}"] = attendee.Company,
        };

        foreach (var (placeholder, value) in replacements)
        {
#pragma warning disable CS0618
            document.ReplaceText(placeholder, value ?? string.Empty);
#pragma warning restore CS0618
        }
    }

    private Task PrintWithWordAsync(string docPath, string printerName, CancellationToken cancellationToken)
    {
        return RunOnStaThreadAsync(() =>
        {
            Word.Application? application = null;
            Word.Document? document = null;

            try
            {
                application = new Word.Application
                {
                    Visible = false
                };

                if (!string.IsNullOrWhiteSpace(printerName))
                {
                    application.ActivePrinter = printerName;
                }

                document = application.Documents.Open(docPath, ReadOnly: true, AddToRecentFiles: false, Visible: false);
                document.PrintOut(Background: false);
            }
            finally
            {
                if (document is not null)
                {
                    document.Close(false);
                    Marshal.FinalReleaseComObject(document);
                }

                if (application is not null)
                {
                    application.Quit(false);
                    Marshal.FinalReleaseComObject(application);
                }
            }
        }, cancellationToken);
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





