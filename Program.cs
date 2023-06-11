using DocumentFormat.OpenXml.Packaging;
using Task3.Classes;

namespace PracticTask3
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Введите путь до файла с данными: ");
            string path = Console.ReadLine();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                int numOfActions = 0;
                while (numOfActions != 4)
                {
                    Console.Clear();
                    Console.WriteLine("Выберите номер действия:");
                    Console.WriteLine("1. По наименованию товара выводить информацию о клиентах, заказавших этот товар, с указанием информации по количеству товара, цене и дате заказа");
                    Console.WriteLine("2. Запрос на изменение контактного лица клиента с указанием параметров: Название организации, ФИО нового контактного лица. В результате информация должна быть занесена в этот же документ, в качестве ответа пользователю необходимо выдавать информацию о результате изменений.");
                    Console.WriteLine("3. Запрос на определение золотого клиента, клиента с наибольшим количеством заказов, за указанный год, месяц.");
                    Console.WriteLine("4. Выход");

                    numOfActions = Int32.Parse(Console.ReadLine());
                    switch (numOfActions)
                    {
                        case 1:
                            OutOfCustomerInformation first = new OutOfCustomerInformation();
                            first.InfoAboutProduct(path, wbPart);

                            break;

                        case 2:
                            ReWriteClients second = new ReWriteClients();
                            second.NewContactNameOfOrganization(document, wbPart);

                            break;

                        case 3:
                            SearchGoldenClients third = new SearchGoldenClients();
                            third.OutGoldenClients(wbPart);

                            break;

                        case 4:
                            Console.Clear();

                            break;

                        default:
                            Console.WriteLine("Некорректный ввод!");

                            break;
                    }
                }
            }
        }
    }
}