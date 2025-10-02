using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Qiosk.App.Models;

namespace Qiosk.App.Services.Contracts;

public interface IQrScannerService : IDisposable
{
    event EventHandler<BitmapSource>? FrameReady;
    event EventHandler<QrCodeDetection>? CodeDetected;

    Task StartAsync(string? cameraMoniker, CancellationToken cancellationToken = default);
    Task StopAsync(CancellationToken cancellationToken = default);
    void RequestPause(TimeSpan duration);
    IReadOnlyList<CameraDevice> GetAvailableCameras();
    CameraControlCapabilities GetCurrentCameraControlCapabilities();
    bool TrySetCameraZoom(int value);
    bool TrySetCameraBrightness(int value);
    bool IsRunning { get; }
}