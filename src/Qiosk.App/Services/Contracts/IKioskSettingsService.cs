using Qiosk.App.Configuration;

namespace Qiosk.App.Services.Contracts;

public interface IKioskSettingsService
{
    ValueTask<KioskSettings> GetAsync(CancellationToken cancellationToken = default);
    ValueTask SaveAsync(KioskSettings settings, CancellationToken cancellationToken = default);
    event EventHandler<KioskSettings>? SettingsChanged;
}
