namespace Qiosk.App.Models;

public sealed record CameraControlCapabilities(CameraControlRange? Zoom, CameraControlRange? Brightness)
{
    public static readonly CameraControlCapabilities Empty = new(null, null);
}