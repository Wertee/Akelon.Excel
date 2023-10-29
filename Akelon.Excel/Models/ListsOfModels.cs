namespace Akelon.Excel.Models;

public record ListsOfModels
{
    public List<Client> Clients { get; set; } = new List<Client>();
    public List<Product> Products { get; set; } = new List<Product>();
    public List<Order> Orders { get; set; } = new List<Order>();
}