using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Runtime.InteropServices;
using DirectShowLib;
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

using AForgeFilterInfoCollection = AForge.Video.DirectShow.FilterInfoCollection;
using AForgeFilterInfo = AForge.Video.DirectShow.FilterInfo;
using AForgeFilterCategory = AForge.Video.DirectShow.FilterCategory;
using CameraControlProperty = DirectShowLib.CameraControlProperty;
using CameraControlFlags = DirectShowLib.CameraControlFlags;
using VideoProcAmpProperty = DirectShowLib.VideoProcAmpProperty;
using VideoProcAmpFlags = DirectShowLib.VideoProcAmpFlags;
namespace Qiosk.App.Services;

public sealed class QrScannerService : IQrScannerService
{
    private readonly IOptionsMonitor<KioskOptions> _optionsMonitor;
    private readonly ILogger<QrScannerService> _logger;
    private readonly ZXing.Windows.Compatibility.BarcodeReader _barcodeReader;
    private readonly SemaphoreSlim _lifecycle = new(1, 1);
    private readonly object _controlSync = new();
    private const int DigitalZoomMinimum = 100;
    private const int DigitalZoomMaximum = 300;
    private const int DigitalZoomStep = 10;
    private const int DigitalZoomDefault = 100;

    private bool _hardwareZoomSupported;
    private bool _hardwareBrightnessSupported;
    private int _digitalZoomValue = DigitalZoomDefault;

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

            lock (_controlSync)
            {
                _digitalZoomValue = DigitalZoomDefault;
                _hardwareZoomSupported = false;
                _hardwareBrightnessSupported = false;
            }

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
            var collection = new AForgeFilterInfoCollection(AForgeFilterCategory.VideoInputDevice);
            return collection.Cast<AForgeFilterInfo>()
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
            using var original = (Bitmap)eventArgs.Frame.Clone();
            Bitmap processing = original;
            Bitmap? zoomed = null;
            try
            {
                bool hardwareZoom;
                int zoomValue;
                lock (_controlSync)
                {
                    hardwareZoom = _hardwareZoomSupported;
                    zoomValue = _digitalZoomValue;
                }

                if (!hardwareZoom)
                {
                    zoomValue = Math.Clamp(zoomValue, DigitalZoomMinimum, DigitalZoomMaximum);
                    if (zoomValue > DigitalZoomMinimum)
                    {
                        var scale = zoomValue / (double)DigitalZoomMinimum;
                        var newWidth = (int)Math.Round(original.Width / scale);
                        var newHeight = (int)Math.Round(original.Height / scale);
                        newWidth = Math.Clamp(newWidth, 1, original.Width);
                        newHeight = Math.Clamp(newHeight, 1, original.Height);
                        var offsetX = (original.Width - newWidth) / 2;
                        var offsetY = (original.Height - newHeight) / 2;
                        var rect = new Rectangle(offsetX, offsetY, newWidth, newHeight);
                        zoomed = original.Clone(rect, original.PixelFormat);
                        processing = zoomed;
                    }
                }

                var frameSource = processing.ToBitmapSource();
                FrameReady?.Invoke(this, frameSource);

                if (DateTime.UtcNow < _pausedUntilUtc)
                {
                    return;
                }

                var result = _barcodeReader.Decode(processing);
                if (result?.Text is { } code && !string.IsNullOrWhiteSpace(code))
                {
                    var cooldown = TimeSpan.FromSeconds(Math.Max(1, _optionsMonitor.CurrentValue.CooldownSeconds));
                    _pausedUntilUtc = DateTime.UtcNow.Add(cooldown);
                    _logger.LogInformation("Cod QR detectat: {Code}", code);
                    CodeDetected?.Invoke(this, new QrCodeDetection(code.Trim(), DateTimeOffset.UtcNow));
                }
            }
            finally
            {
                zoomed?.Dispose();
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
            lock (_controlSync)
            {
                _hardwareZoomSupported = false;
                _hardwareBrightnessSupported = false;
                _digitalZoomValue = DigitalZoomDefault;
            }

            _device = null;
            IsRunning = false;
        }
    }


    public CameraControlCapabilities GetCurrentCameraControlCapabilities()
    {
        ThrowIfDisposed();

        if (_device is null)
        {
            return CameraControlCapabilities.Empty;
        }

        lock (_controlSync)
        {
            var zoomRange = TryGetCameraControlRange(CameraControlProperty.Zoom);
            _hardwareZoomSupported = zoomRange is not null;

            if (!_hardwareZoomSupported)
            {
                var clampedZoom = Math.Clamp(_digitalZoomValue, DigitalZoomMinimum, DigitalZoomMaximum);
                _digitalZoomValue = clampedZoom;
                zoomRange = new CameraControlRange(DigitalZoomMinimum, DigitalZoomMaximum, DigitalZoomStep, DigitalZoomDefault, clampedZoom);
            }
            else
            {
                _digitalZoomValue = DigitalZoomDefault;
            }

            var brightnessRange = TryGetVideoProcAmpRange(VideoProcAmpProperty.Brightness);
            _hardwareBrightnessSupported = brightnessRange is not null;

            return new CameraControlCapabilities(zoomRange, brightnessRange);
        }
    }

    public bool TrySetCameraZoom(int value)
    {
        ThrowIfDisposed();

        if (_device is null)
        {
            return false;
        }

        lock (_controlSync)
        {
            if (_hardwareZoomSupported)
            {
                return TrySetCameraControlProperty(CameraControlProperty.Zoom, value);
            }

            var clamped = Math.Clamp(value, DigitalZoomMinimum, DigitalZoomMaximum);
            _digitalZoomValue = clamped;
            return true;
        }
    }

    public bool TrySetCameraBrightness(int value)
    {
        ThrowIfDisposed();

        if (_device is null)
        {
            return false;
        }

        lock (_controlSync)
        {
            if (!_hardwareBrightnessSupported)
            {
                return false;
            }

            return TrySetVideoProcAmpProperty(VideoProcAmpProperty.Brightness, value);
        }
    }

    private CameraControlRange? TryGetCameraControlRange(CameraControlProperty property)
    {
        if (_device is null)
        {
            return null;
        }

        object? source = null;
        try
        {
            source = _device.SourceObject;
            if (source is not IAMCameraControl cameraControl)
            {
                return null;
            }

            var hr = cameraControl.GetRange(property, out int min, out int max, out int step, out int @default, out CameraControlFlags _);
            if (hr < 0)
            {
                return null;
            }

            hr = cameraControl.Get(property, out int current, out CameraControlFlags _);
            if (hr < 0)
            {
                current = @default;
            }

            return new CameraControlRange(min, max, Math.Max(1, step), @default, current);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Nu am putut obtine intervalul pentru {Property}", property);
            return null;
        }
        finally
        {
            if (source is not null)
            {
                Marshal.ReleaseComObject(source);
            }
        }
    }

    private CameraControlRange? TryGetVideoProcAmpRange(VideoProcAmpProperty property)
    {
        if (_device is null)
        {
            return null;
        }

        object? source = null;
        try
        {
            source = _device.SourceObject;
            if (source is not IAMVideoProcAmp videoProcAmp)
            {
                return null;
            }

            var hr = videoProcAmp.GetRange(property, out int min, out int max, out int step, out int @default, out VideoProcAmpFlags _);
            if (hr < 0)
            {
                return null;
            }

            hr = videoProcAmp.Get(property, out int current, out VideoProcAmpFlags _);
            if (hr < 0)
            {
                current = @default;
            }

            return new CameraControlRange(min, max, Math.Max(1, step), @default, current);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Nu am putut obtine intervalul video pentru {Property}", property);
            return null;
        }
        finally
        {
            if (source is not null)
            {
                Marshal.ReleaseComObject(source);
            }
        }
    }

    private bool TrySetCameraControlProperty(CameraControlProperty property, int value)
    {
        object? source = null;
        try
        {
            source = _device?.SourceObject;
            if (source is not IAMCameraControl cameraControl)
            {
                return false;
            }

            var hr = cameraControl.Set(property, value, CameraControlFlags.Manual);
            return hr >= 0;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Nu am putut seta {Property} pentru camera.", property);
            return false;
        }
        finally
        {
            if (source is not null)
            {
                Marshal.ReleaseComObject(source);
            }
        }
    }

    private bool TrySetVideoProcAmpProperty(VideoProcAmpProperty property, int value)
    {
        object? source = null;
        try
        {
            source = _device?.SourceObject;
            if (source is not IAMVideoProcAmp videoProcAmp)
            {
                return false;
            }

            var hr = videoProcAmp.Set(property, value, VideoProcAmpFlags.Manual);
            return hr >= 0;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Nu am putut seta proprietatea video {Property}.", property);
            return false;
        }
        finally
        {
            if (source is not null)
            {
                Marshal.ReleaseComObject(source);
            }
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
