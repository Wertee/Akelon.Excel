using Akelon.Excel.DataService;

public class Program
{
    private static DataService _dataService;
    public static void Main()
    {
        Console.WriteLine("Укажите путь до файла:");
        string? path = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(path))
        {
            Console.WriteLine("Вы не указали путь до файла");
            return;
        }

        if (!File.Exists(path))
        {
            Console.WriteLine("По заданному пути файл не найден");
            return;
        }
        
        _dataService = new DataService(path);
        ShowUserMenu();

        Console.ReadKey();

    }

    public static void ShowUserMenu()
    {
        Console.Clear();
            
        Console.WriteLine("Выберите операцию:\n" +
                          "1. Получить информацию о клиентах заказавших товар\n" +
                          "2. Изменить контактное лицо для организации\n" +
                          "3. Узнать золотого клиента\n" +
                          "4. Выйти из программы");
        bool correctNumber = int.TryParse(Console.ReadLine(), out int taskNumber);
        if (!correctNumber || taskNumber < 1 || taskNumber > 4)
        {
            Console.WriteLine("Номер операции должен быть числом от 1 до 4");
            ShowUserMenu();
        }
        switch (taskNumber)
        {
            case 1:
                MenuTaskGetClientInfoByProductNameFromOrders();
                break;
            case 2:
                MenuTaskUpdateClientInfo();
                break;
            case 3:
                MenuTaskFindGoldenClientFromOrders();
                break;
            case 4:
                Environment.Exit(0);
                break;
        }
    }

    public static void MenuTaskGetClientInfoByProductNameFromOrders()
    {
        Console.Clear();
        Console.WriteLine("Введите наименование интересующего товара: ");
        string productName = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(productName))
        {
            Console.WriteLine("Вы не ввели название товара\n" +
                              "Для повторного ввода введите 1\n" +
                              "Для выхода в главное меню введите любой символ");
            bool parseResult = int.TryParse(Console.ReadLine(), out int num);
            if(!parseResult)
                ShowUserMenu();
            if(num == 1)
                MenuTaskGetClientInfoByProductNameFromOrders();
            else
                ShowUserMenu();

        }
        
        _dataService.GetClientInfoByProductNameFromOrders(productName);
        
        Console.WriteLine("Для возврата в главное меню нажмите любую кнопку");
        Console.ReadKey();
        ShowUserMenu();
    }

    public static void MenuTaskUpdateClientInfo()
    {
        Console.Clear();
        Console.WriteLine("Введите полное наименование организации: ");
        string? organisationName = Console.ReadLine();
        Console.WriteLine("Введите новое контактное лицо: ");
        string? newContactName = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(organisationName) || string.IsNullOrWhiteSpace(newContactName))
        {
            Console.WriteLine("Вы не ввели название организации или контактное лицо\n" +
                              "Для повторного ввода введите 1\n" +
                              "Для выхода в главное меню введите любой символ");
            bool parseResult = int.TryParse(Console.ReadLine(), out int num);
            if(!parseResult)
                ShowUserMenu();
            if(num == 1)
                MenuTaskUpdateClientInfo();
            else
                ShowUserMenu();
        }
        
        _dataService.UpdateClientInfo(organisationName,newContactName);
        Console.WriteLine("Для возврата в главное меню нажмите любую кнопку");
        Console.ReadKey();
        ShowUserMenu();
    }

    public static void MenuTaskFindGoldenClientFromOrders()
    {
        Console.Clear();
        Console.WriteLine("Введите интересующий вас год: ");
        bool isYearCorrect = int.TryParse(Console.ReadLine(), out int year);
        Console.WriteLine("Введите номер месяца: ");
        bool isMonthCorrect = int.TryParse(Console.ReadLine(), out int month);
        if (!isYearCorrect || !isMonthCorrect || month > 12 || month < 1)
        {
            Console.WriteLine("Вы ввели некорректный год или месяц\n" +
                              "Для повторного ввода введите 1\n" +
                              "Для выхода в главное меню введите любой символ");
            bool parseResult = int.TryParse(Console.ReadLine(), out int num);
            if(!parseResult)
                ShowUserMenu();
            if(num == 1)
                MenuTaskFindGoldenClientFromOrders();
            else
                ShowUserMenu();
        }
        
        _dataService.FindGoldenClientFromOrders(year,month);
        Console.WriteLine("Для возврата в главное меню нажмите любую кнопку");
        Console.ReadKey();
        ShowUserMenu();
    }
    
}
