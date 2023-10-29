using Akelon.Excel.Models;
using ClosedXML.Excel;

namespace Akelon.Excel.DataService;

public class DataService
{
    private readonly string _path;
    public DataService(string path)
    {
        _path = path;
    }
    
    public void GetClientInfoByProductNameFromOrders(string productName)
    {
        ModelsFiller modelsFiller = new();
        ListsOfModels listsOfModelsList = new ListsOfModels();
        try
        {
            listsOfModelsList = modelsFiller.FillListOfModels(_path);
        }
        catch (Exception exception)
        {
            Console.WriteLine("Ошибка считывания данных из файла. Возможно вы выбрали некорректный файл");
            return;
        }
        
        var result = from order in listsOfModelsList.Orders
            join product in listsOfModelsList.Products on order.ProductId equals product.ProductId
            join client in listsOfModelsList.Clients on order.ClientId equals client.ClientId
            where string.Equals(product.Name, productName, StringComparison.CurrentCultureIgnoreCase)
            select new
            {
                orderID = order.OrderId,
                productName = product.Name,
                contactName = client.ContactName,
                clientOrganisation = client.OrganisationName,
                productCount = order.Count,
                productPrice = product.Price,
                orderDateOfCreation = order.DateOfCreation
            };
        if (result.Any())
        {
            foreach (var item in result)
            {
                Console.Clear();
                Console.WriteLine($"Поиск по заказу товара: {item.productName}\n"+
                                  $"Клиент: {item.clientOrganisation}\n" + 
                                  $"Контактное лицо: {item.contactName}\n"+
                                  $"Количество заказанного товара: {item.productCount}\n"+
                                  $"Цена за единицу товара: {item.productPrice}\n"+
                                  $"Дата создания заказа: {item.orderDateOfCreation}");
            
            }
        }
        else
        {
            Console.WriteLine($"Данных по продукту {productName} не найдено!");
        }
        
    }

    public void UpdateClientInfo(string organisationName,string newContactName)
    {
        using var workBook = new XLWorkbook(_path);
        IEnumerable<IXLRangeRow>? rows;
        try
        {
            rows = workBook.Worksheet(2).RangeUsed().RowsUsed().Skip(1);
        }
        catch (Exception exception)
        {
            Console.WriteLine("Ошибка получения необходимого листа в XLSX файле. Возможно вы выбрали некорректный файл");
            return;
        }
        bool isFound = false;
        foreach (var row in rows)
        {
            if (string.Equals(row.Cell(2).GetString(), organisationName, StringComparison.CurrentCultureIgnoreCase))
            {
                row.Cell(4).Value = newContactName;
                isFound = true;
            }
        }
        Console.Clear();
        try
        {
            workBook.Save();
        }
        catch (Exception exception)
        {
            Console.WriteLine($"Ошибка при сохранении файла: {exception.Message}");
        }
        
        if(isFound)
            Console.WriteLine($"Контактное лицо у организации {organisationName} изменено на {newContactName}");
        else
            Console.WriteLine($"Организация {organisationName} не найдена!");
    }

    public void FindGoldenClientFromOrders(int year, int month)
    {
        if (month < 1 || month > 12)
        {
            Console.WriteLine("Месяц должен быть указан в интервале 1-12");
            return;
        }
        
        string[] months = { "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };
        ModelsFiller modelsFiller = new();
        ListsOfModels listsOfModelsList = new ListsOfModels();
        try
        {
            listsOfModelsList = modelsFiller.FillListOfModels(_path);
        }
        catch (Exception exception)
        {
            Console.WriteLine("Ошибка считывания данных из файла. Возможно вы выбрали некорректный файл");
        }
        var result = from order in listsOfModelsList.Orders
            join product in listsOfModelsList.Products on order.ProductId equals product.ProductId
            join client in listsOfModelsList.Clients on order.ClientId equals client.ClientId
            where order.DateOfCreation.Month == month && order.DateOfCreation.Year == year
            group order by client.OrganisationName into ord
            select new
            {
                organisationName = ord.Key,
                countOfOrders = ord.Count()
            };
        if (result.Any())
        {
            Console.Clear();
            if(result.Count() > 1)
                Console.WriteLine("В указанном вами месяце несколько золотых клиентов");
            foreach (var item in result)
            {
                Console.WriteLine("========================");
                Console.WriteLine($"Золотой клиент за {months[month-1]}\n"+
                                  $"Организация: {item.organisationName}\n"+
                                  $"Количество заказов {item.countOfOrders}");
            }
        }
        else
        {
            Console.WriteLine("По выбранному месяцу и году заказов не найдено");
        }
    }


}