namespace Akelon.Excel.Models;

public record Order
{
    public int OrderId { get; init; }
    public int ProductId { get; init; }
    public int ClientId { get; init; }
    public int OrderNumber { get; init; }
    public int Count { get; init; }
    public DateOnly DateOfCreation { get; init; }
}