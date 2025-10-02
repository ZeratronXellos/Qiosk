using System;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Qiosk.App.Configuration;
using Qiosk.App.Services.Contracts;

namespace Qiosk.App.Services;

public sealed class KioskSettingsService : IKioskSettingsService, IDisposable
{
    private readonly string _settingsPath;
    private readonly SemaphoreSlim _sync = new(1, 1);
    private readonly JsonSerializerOptions _serializerOptions = new()
    {
        WriteIndented = true
    };

    private KioskSettings? _cache;
    private bool _disposed;

    public KioskSettingsService()
    {
        _settingsPath = Path.Combine(AppContext.BaseDirectory, "kiosk-settings.json");
    }

    public event EventHandler<KioskSettings>? SettingsChanged;

    public async ValueTask<KioskSettings> GetAsync(CancellationToken cancellationToken = default)
    {
        await EnsureNotDisposedAsync();
        await _sync.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_cache is null)
            {
                _cache = await ReadFromDiskAsync(cancellationToken).ConfigureAwait(false);
            }

            return Clone(_cache);
        }
        finally
        {
            _sync.Release();
        }
    }

    public async ValueTask SaveAsync(KioskSettings settings, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(settings);
        await EnsureNotDisposedAsync();

        await _sync.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            _cache = Clone(settings);
            await WriteToDiskAsync(_cache, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _sync.Release();
        }

        SettingsChanged?.Invoke(this, Clone(settings));
    }

    private async Task<KioskSettings> ReadFromDiskAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_settingsPath))
        {
            var defaults = new KioskSettings();
            await WriteToDiskAsync(defaults, cancellationToken).ConfigureAwait(false);
            return defaults;
        }

        await using var stream = File.OpenRead(_settingsPath);
        var settings = await JsonSerializer.DeserializeAsync<KioskSettings>(stream, _serializerOptions, cancellationToken).ConfigureAwait(false);
        return settings ?? new KioskSettings();
    }

    private async Task WriteToDiskAsync(KioskSettings settings, CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(_settingsPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        await using var stream = File.Create(_settingsPath);
        await JsonSerializer.SerializeAsync(stream, settings, _serializerOptions, cancellationToken).ConfigureAwait(false);
    }

    private static KioskSettings Clone(KioskSettings settings) => new()
    {
        DefaultPrinter = settings.DefaultPrinter,
        DefaultCameraMoniker = settings.DefaultCameraMoniker,
        LockMode = settings.LockMode,
        ExcelPathOverride = settings.ExcelPathOverride,
        TemplatePathOverride = settings.TemplatePathOverride,
        ReloadOnStart = settings.ReloadOnStart
    };

    private Task EnsureNotDisposedAsync()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(KioskSettingsService));
        }

        return Task.CompletedTask;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _sync.Dispose();
        _disposed = true;
    }
}
