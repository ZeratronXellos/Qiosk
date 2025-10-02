using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using AForge.Video;
using AForge.Video.DirectShow;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Qiosk.App.Configuration;
using Qiosk.App.Infrastructure;
using Qiosk.App.Models;
using Qiosk.App.Services.Contracts;
using ZXing;
using ZXing.Common;

namespace Qiosk.App.Services;

public sealed class QrScannerService : IQrScannerService
{
    private readonly IOptionsMonitor<KioskOptions> _optionsMonitor;
    private readonly ILogger<QrScannerService> _logger;
    private readonly ZXing.Windows.Compatibility.BarcodeReader _barcodeReader;
    private readonly SemaphoreSlim _lifecycle = new(1, 1);

    private VideoCaptureDevice? _device;
    private bool _disposed;
    private DateTime _pausedUntilUtc = DateTime.MinValue;

    public QrScannerService(IOptionsMonitor<KioskOptions> optionsMonitor, ILogger<QrScannerService> logger)
    {
        _optionsMonitor = optionsMonitor;
        _logger = logger;
        _barcodeReader = new ZXing.Windows.Compatibility.BarcodeReader
        {
            AutoRotate = true,
            Options = new DecodingOptions
            {
                PossibleFormats = new List<BarcodeFormat> { BarcodeFormat.QR_CODE },
                TryHarder = true,
                TryInverted = true
            }
        };
    }

    public event EventHandler<BitmapSource>? FrameReady;
    public event EventHandler<QrCodeDetection>? CodeDetected;

    public bool IsRunning { get; private set; }

    public async Task StartAsync(string? cameraMoniker, CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        await _lifecycle.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await StopInternalAsync().ConfigureAwait(false);

            var moniker = cameraMoniker;
            if (string.IsNullOrWhiteSpace(moniker))
            {
                moniker = GetAvailableCameras().FirstOrDefault()?.MonikerString;
            }

            if (string.IsNullOrWhiteSpace(moniker))
            {
                throw new InvalidOperationException("Nu exista camere disponibile pentru scanare.");
            }

            _device = new VideoCaptureDevice(moniker);
            _device.NewFrame += HandleNewFrame;
            _device.PlayingFinished += HandlePlayingFinished;

            _logger.LogInformation("Porneste camera cu moniker {Moniker}", moniker);
            _device.Start();
            IsRunning = true;
        }
        finally
        {
            _lifecycle.Release();
        }
    }

    public async Task StopAsync(CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        await _lifecycle.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await StopInternalAsync().ConfigureAwait(false);
        }
        finally
        {
            _lifecycle.Release();
        }
    }

    public void RequestPause(TimeSpan duration)
    {
        var requestedUntil = DateTime.UtcNow.Add(duration);
        if (requestedUntil > _pausedUntilUtc)
        {
            _pausedUntilUtc = requestedUntil;
        }
    }

    public IReadOnlyList<CameraDevice> GetAvailableCameras()
    {
        ThrowIfDisposed();
        try
        {
            var collection = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            return collection.Cast<FilterInfo>()
                .Select(info => new CameraDevice(info.Name, info.MonikerString))
                .ToList();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Nu am putut enumera camerele disponibile.");
            return Array.Empty<CameraDevice>();
        }
    }

    private void HandleNewFrame(object? sender, NewFrameEventArgs eventArgs)
    {
        if (_disposed)
        {
            return;
        }

        try
        {
            using var bitmap = (Bitmap)eventArgs.Frame.Clone();
            var frameSource = bitmap.ToBitmapSource();
            FrameReady?.Invoke(this, frameSource);

            if (DateTime.UtcNow < _pausedUntilUtc)
            {
                return;
            }

            var result = _barcodeReader.Decode(bitmap);
            if (result?.Text is { } code && !string.IsNullOrWhiteSpace(code))
            {
                var cooldown = TimeSpan.FromSeconds(Math.Max(1, _optionsMonitor.CurrentValue.CooldownSeconds));
                _pausedUntilUtc = DateTime.UtcNow.Add(cooldown);
                _logger.LogInformation("Cod QR detectat: {Code}", code);
                CodeDetected?.Invoke(this, new QrCodeDetection(code.Trim(), DateTimeOffset.UtcNow));
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Eroare la procesarea frame-ului video.");
        }
    }

    private void HandlePlayingFinished(object? sender, ReasonToFinishPlaying reason)
    {
        _logger.LogWarning("Camera s-a oprit cu motivul {Reason}", reason);
        IsRunning = false;
    }

    private async Task StopInternalAsync()
    {
        if (_device is null)
        {
            IsRunning = false;
            return;
        }

        try
        {
            _logger.LogInformation("Oprire camera QR");
            _device.NewFrame -= HandleNewFrame;
            _device.PlayingFinished -= HandlePlayingFinished;
            if (_device.IsRunning)
            {
                _device.SignalToStop();
                await Task.Run(() => _device.WaitForStop());
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Eroare la oprirea camerei");
        }
        finally
        {
            _device = null;
            IsRunning = false;
        }
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(QrScannerService));
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        StopInternalAsync().GetAwaiter().GetResult();
        _lifecycle.Dispose();
    }
}
