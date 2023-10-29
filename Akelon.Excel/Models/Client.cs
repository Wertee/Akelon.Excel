namespace Akelon.Excel.Models;

public record Client
{
    public int ClientId { get; init; }
    public string OrganisationName { get; init; }
    public string Address { get; init; }
    public string ContactName { get; init; }
}