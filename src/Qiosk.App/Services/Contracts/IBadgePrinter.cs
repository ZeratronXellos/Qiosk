using System.Threading;
using System.Threading.Tasks;
using Qiosk.App.Models;

namespace Qiosk.App.Services.Contracts;

public interface IBadgePrinter
{
    Task PrintAsync(Attendee attendee, string printerName, string templatePath, CancellationToken cancellationToken = default);
}
