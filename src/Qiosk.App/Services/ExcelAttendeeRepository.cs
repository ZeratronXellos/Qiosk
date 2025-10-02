using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Qiosk.App.Configuration;
using Qiosk.App.Models;
using Qiosk.App.Services.Contracts;

namespace Qiosk.App.Services;

public sealed class ExcelAttendeeRepository : IAttendeeRepository, IDisposable
{
    private readonly IOptionsMonitor<KioskOptions> _optionsMonitor;
    private readonly IKioskSettingsService _settingsService;
    private readonly ILogger<ExcelAttendeeRepository> _logger;
    private readonly SemaphoreSlim _loadSync = new(1, 1);
    private readonly SemaphoreSlim _writeSync = new(1, 1);
    private const int IdColumn = 1;
    private const int LastNameColumn = 2;
    private const int FirstNameColumn = 3;
    private const int RoleColumn = 4;
    private const int CompanyColumn = 5;
    private const int StatusColumn = 6;
    private const int CheckInColumn = 7;
    private const int BadgeColumn = 8;

    private readonly Dictionary<string, Attendee> _cache = new(StringComparer.OrdinalIgnoreCase);

    private bool _disposed;

    public ExcelAttendeeRepository(
        IOptionsMonitor<KioskOptions> optionsMonitor,
        IKioskSettingsService settingsService,
        ILogger<ExcelAttendeeRepository> logger)
    {
        _optionsMonitor = optionsMonitor;
        _settingsService = settingsService;
        _logger = logger;
    }

    public IReadOnlyCollection<Attendee> All
    {
        get
        {
            lock (_cache)
            {
                return _cache.Values.ToList().AsReadOnly();
            }
        }
    }

    public Attendee? FindById(string id)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(id);
        lock (_cache)
        {
            return _cache.TryGetValue(id.Trim(), out var attendee) ? attendee : null;
        }
    }

    public async ValueTask LoadAsync(CancellationToken cancellationToken = default)
    {
        await EnsureNotDisposedAsync();
        await _loadSync.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var settings = await _settingsService.GetAsync(cancellationToken).ConfigureAwait(false);
            var options = _optionsMonitor.CurrentValue;

            var excelPath = ResolvePath(settings.ExcelPathOverride, options.ExcelPath);

            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException($"Nu gasesc fisierul Excel la calea '{excelPath}'.", excelPath);
            }

            _logger.LogInformation("Loading attendees from {Path}", excelPath);

            using var workbook = new XLWorkbook(excelPath);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            if (worksheet is null)
            {
                _logger.LogWarning("Fisierul Excel {Path} nu contine foi de lucru.", excelPath);
                lock (_cache)
                {
                    _cache.Clear();
                }
                return;
            }

            var rows = worksheet.RangeUsed()?.RowsUsed().Skip(1) ?? Enumerable.Empty<IXLRangeRow>();
            var items = new Dictionary<string, Attendee>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in rows)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var id = row.Cell(IdColumn).GetValue<string>().Trim();
                if (string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }

                var lastName = row.Cell(LastNameColumn).GetValue<string>().Trim();
                var firstName = row.Cell(FirstNameColumn).GetValue<string>().Trim();
                var role = row.Cell(RoleColumn).GetValue<string>().Trim();
                var company = row.Cell(CompanyColumn).GetValue<string>().Trim();

                var attendee = new Attendee(id, lastName, firstName, role, company);
                items[id] = attendee;
            }

            lock (_cache)
            {
                _cache.Clear();
                foreach (var kvp in items)
                {
                    _cache.Add(kvp.Key, kvp.Value);
                }
            }

            _logger.LogInformation("Loaded {Count} attendees from Excel.", items.Count);
        }
        finally
        {
            _loadSync.Release();
        }
    }

    public async ValueTask MarkPresentAsync(string attendeeId, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(attendeeId);
        await EnsureNotDisposedAsync().ConfigureAwait(false);

        var normalizedId = attendeeId.Trim();
        var settings = await _settingsService.GetAsync(cancellationToken).ConfigureAwait(false);
        var options = _optionsMonitor.CurrentValue;
        var excelPath = ResolvePath(settings.ExcelPathOverride, options.ExcelPath);

        if (!File.Exists(excelPath))
        {
            throw new FileNotFoundException($"Nu gasesc fisierul Excel la calea '{excelPath}'.", excelPath);
        }

        await _writeSync.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await Task.Run(() =>
            {
                using var workbook = new XLWorkbook(excelPath);
                var worksheet = workbook.Worksheets.FirstOrDefault();
                if (worksheet is null)
                {
                    _logger.LogWarning("Fisierul Excel {Path} nu contine foi de lucru.", excelPath);
                    return;
                }

                var rows = worksheet.RangeUsed()?.RowsUsed().Skip(1) ?? Enumerable.Empty<IXLRangeRow>();
                var updated = false;

                foreach (var row in rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var id = row.Cell(IdColumn).GetValue<string>().Trim();
                    if (!string.Equals(id, normalizedId, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var statusCell = row.Cell(StatusColumn);
                    var currentValue = statusCell.GetValue<string>();
                    if (string.IsNullOrWhiteSpace(currentValue))
                    {
                        statusCell.Value = "Prezent";
                        var checkInCell = row.Cell(CheckInColumn);
                        if (string.IsNullOrWhiteSpace(checkInCell.GetValue<string>()))
                        {
                            checkInCell.Value = DateTime.Now.ToString("HH:mm", CultureInfo.InvariantCulture);
                        }

                        updated = true;
                    }

                    break;
                }

                if (updated)
                {
                    workbook.Save();
                }
            }, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _writeSync.Release();
        }
    }

    public async ValueTask MarkBadgePrintedAsync(string attendeeId, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(attendeeId);
        await EnsureNotDisposedAsync().ConfigureAwait(false);

        var normalizedId = attendeeId.Trim();
        var settings = await _settingsService.GetAsync(cancellationToken).ConfigureAwait(false);
        var options = _optionsMonitor.CurrentValue;
        var excelPath = ResolvePath(settings.ExcelPathOverride, options.ExcelPath);

        if (!File.Exists(excelPath))
        {
            throw new FileNotFoundException($"Nu gasesc fisierul Excel la calea '{excelPath}'.", excelPath);
        }

        await _writeSync.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await Task.Run(() =>
            {
                using var workbook = new XLWorkbook(excelPath);
                var worksheet = workbook.Worksheets.FirstOrDefault();
                if (worksheet is null)
                {
                    _logger.LogWarning("Fisierul Excel {Path} nu contine foi de lucru.", excelPath);
                    return;
                }

                var rows = worksheet.RangeUsed()?.RowsUsed().Skip(1) ?? Enumerable.Empty<IXLRangeRow>();
                var updated = false;

                foreach (var row in rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var id = row.Cell(IdColumn).GetValue<string>().Trim();
                    if (!string.Equals(id, normalizedId, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var badgeCell = row.Cell(BadgeColumn);
                    if (string.IsNullOrWhiteSpace(badgeCell.GetValue<string>()))
                    {
                        badgeCell.Value = "DA";
                        updated = true;
                    }

                    break;
                }

                if (updated)
                {
                    workbook.Save();
                }
            }, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _writeSync.Release();
        }
    }

    private static string ResolvePath(string overridePath, string defaultPath)
    {
        var candidate = string.IsNullOrWhiteSpace(overridePath) ? defaultPath : overridePath;
        if (string.IsNullOrWhiteSpace(candidate))
        {
            throw new InvalidOperationException("Nu este configurata calea catre fi?ierul Excel.");
        }

        if (!Path.IsPathRooted(candidate))
        {
            candidate = Path.Combine(AppContext.BaseDirectory, candidate);
        }

        return Path.GetFullPath(candidate);
    }

    private Task EnsureNotDisposedAsync()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(ExcelAttendeeRepository));
        }

        return Task.CompletedTask;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _loadSync.Dispose();
        _writeSync.Dispose();
        _disposed = true;
    }
}
