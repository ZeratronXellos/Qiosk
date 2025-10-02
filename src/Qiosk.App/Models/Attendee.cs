using System.Linq;

namespace Qiosk.App.Models;

public sealed record Attendee(
    string Id,
    string LastName,
    string FirstName,
    string Role,
    string Company)
{
    public string FullName => string.Join(" ", new[] { FirstName, LastName }.Where(p => !string.IsNullOrWhiteSpace(p)));
}
