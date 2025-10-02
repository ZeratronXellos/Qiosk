using System.Threading;
using System.Threading.Tasks;
using Qiosk.App.Models;

namespace Qiosk.App.Services.Contracts;

public interface IAttendeeRepository
{
    ValueTask LoadAsync(CancellationToken cancellationToken = default);
    Attendee? FindById(string id);
    IReadOnlyCollection<Attendee> All { get; }
    ValueTask MarkPresentAsync(string attendeeId, CancellationToken cancellationToken = default);
}