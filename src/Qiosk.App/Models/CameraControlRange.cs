namespace Qiosk.App.Models;

public sealed record CameraControlRange(int Minimum, int Maximum, int Step, int Default, int Current);