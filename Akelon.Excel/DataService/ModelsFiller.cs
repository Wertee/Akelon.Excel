using Akelon.Excel.Models;
using ClosedXML.Excel;

namespace Akelon.Excel.DataService;

public class ModelsFiller
{
    public ListsOfModels FillListOfModels(string path)
    {
        using var workBook = new XLWorkbook(path);
        return new ListsOfModels()
        {
            Products = GetProducts(workBook),
            Clients = GetClients(workBook),
            Orders = GetOrders(workBook)
        };
    }

    private List<Product> GetProducts(IXLWorkbook workbook)
    {
        var rows = workbook.Worksheet(1).RangeUsed().RowsUsed().Skip(1);
        List<Product> productList = new();
        foreach (var row in rows)
        {
            Product product = new()
            {
                ProductId = int.Parse(row.Cell(1).GetString()),
                Name = row.Cell(2).GetString(),
                Unit = row.Cell(3).GetString(),
                Price = decimal.Parse(row.Cell(4).GetString())
            };
            productList.Add(product);
        }

        return productList;
    }

    private List<Order> GetOrders(IXLWorkbook workbook)
    {
        var rows = workbook.Worksheet(3).RangeUsed().RowsUsed().Skip(1);
        List<Order> orders = new();
        foreach (var row in rows)
        {
            Order order = new()
            {
                OrderId = int.Parse(row.Cell(1).GetString()),
                ProductId = int.Parse(row.Cell(2).GetString()),
                ClientId = int.Parse(row.Cell(3).GetString()),
                OrderNumber = int.Parse(row.Cell(4).GetString()),
                Count = int.Parse(row.Cell(5).GetString()),
                DateOfCreation = DateOnly.FromDateTime(DateTime.Parse(row.Cell(6).GetString()))
            };
            orders.Add(order);
        }

        return orders;
    }

    private List<Client> GetClients(IXLWorkbook workbook)
    {
        var rows = workbook.Worksheet(2).RangeUsed().RowsUsed().Skip(1);
        List<Client> clientList = new();
        foreach (var row in rows)
        {
            Client client = new()
            {
                ClientId = int.Parse(row.Cell(1).GetString()),
                OrganisationName = row.Cell(2).GetString(),
                Address = row.Cell(3).GetString(),
                ContactName = row.Cell(4).GetString()
            };
            clientList.Add(client);
        }

        return clientList;
    }
}