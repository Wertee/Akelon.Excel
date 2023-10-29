using Akelon.Excel.DataService;

public class Program
{
    private static DataService _dataService;
    public static void Main()
    {
        Console.Write("Укажите путь до файла:");
        string? path = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            Console.WriteLine("Вы не указали путь до файла");
            Main();
        }
        
        _dataService = new DataService(path);
        ShowUserMenu();

        Console.ReadKey();

    }

    private static void ShowUserMenu()
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
            Console.WriteLine("Номер операции должен быть числом от 1 до 4\nДля возврата в главное меню нажмите любую кнопку");
            Console.ReadKey();
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

    private static void MenuTaskGetClientInfoByProductNameFromOrders()
    {
        Console.Clear();
        Console.Write("Введите наименование интересующего товара: ");
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

    private static void MenuTaskUpdateClientInfo()
    {
        Console.Clear();
        Console.Write("Введите полное наименование организации: ");
        string? organisationName = Console.ReadLine();
        Console.Write("Введите новое контактное лицо: ");
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

    private static void MenuTaskFindGoldenClientFromOrders()
    {
        Console.Clear();
        Console.Write("Введите интересующий вас год: ");
        bool isYearCorrect = int.TryParse(Console.ReadLine(), out int year);
        Console.Write("Введите номер месяца: ");
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
