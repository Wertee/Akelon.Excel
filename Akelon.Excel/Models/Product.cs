namespace Akelon.Excel.Models;

public record Product
{
    public int ProductId { get; init; }
    public string Name { get; init; }
    public string Unit { get; init; }
    public decimal Price { get; init; }
    
}