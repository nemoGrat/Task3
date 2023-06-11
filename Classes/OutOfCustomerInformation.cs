using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

namespace Task3.Classes
{
    internal class OutOfCustomerInformation
    {
        string inputProductName;
        string productName;
        string productPrice;
        string clientCode;
        string orderVolume;
        string nameOfOrganization;
        DateTime orderDate;

        private int lineNumber;
        private bool flag;
        private string value;

        public OutOfCustomerInformation()
        {
            inputProductName = productName = productPrice = clientCode = orderVolume = nameOfOrganization = null;
        }

        public string ReadCell(string sheetName, string cellName, WorkbookPart wbPart)
        {
            string value = null;

            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellName).FirstOrDefault();

            if (theCell.InnerText.Length > 0)
            {
                value = theCell.InnerText;
                if (theCell.DataType != null)
                {
                    var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (stringTable != null) value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }

            return value;
        }

        public void InputNameOfProduct()
        {
            Console.Write("Введите наименование товара: ");
            inputProductName = Console.ReadLine();
        }

        public void SearchNameAndPrice(string path, WorkbookPart wbPart)
        {
            flag = false;
            lineNumber = 2;
            value = null;

            do
            {
                value = ReadCell("Товары", "B" + lineNumber.ToString(), wbPart);

                if (value == inputProductName)
                {
                    productName = ReadCell("Товары", "A" + lineNumber.ToString(), wbPart);
                    productPrice = ReadCell("Товары", "D" + lineNumber.ToString(), wbPart);

                    flag = true;
                }
                else lineNumber++;

            } while (value != null && !flag);
        }

        public void SearchCodeAndVolumeAndDate(string path, WorkbookPart wbPart)
        {
            flag = false;
            lineNumber = 2;
            value = null;

            do
            {
                value = ReadCell("Заявки", "B" + lineNumber.ToString(), wbPart);

                if (value == productName)
                {
                    clientCode = ReadCell("Заявки", "C" + lineNumber.ToString(), wbPart);
                    orderVolume = ReadCell("Заявки", "E" + lineNumber.ToString(), wbPart);
                    orderDate = new DateTime(1900, 1, 1, 0, 0, 0).AddDays(Int32.Parse(ReadCell("Заявки", "F" + lineNumber.ToString(), wbPart)) - 2);

                    flag = true;
                }
                else lineNumber++;

            } while (value != null && !flag);
        }

        public void SearchNameOfOrganization(string path, WorkbookPart wbPart)
        {
            flag = false;
            lineNumber = 2;
            value = null;

            do
            {
                value = ReadCell("Клиенты", "A" + lineNumber.ToString(), wbPart);

                if (value == clientCode)
                {
                    nameOfOrganization = ReadCell("Клиенты", "B" + lineNumber.ToString(), wbPart);

                    flag = true;
                }
                else lineNumber++;

            } while (value != null && !flag);
        }

        public void OutInfoAboutProduct()
        {
            if (orderVolume != null)
            {
                Console.WriteLine("Товар \"" + inputProductName + "\" заказала организация:");
                Console.WriteLine(nameOfOrganization + ", " + orderDate.Day + "." + orderDate.Month + "." + orderDate.Year
                    + ", в количестве " + orderVolume + " штук, на сумму "
                    + Int32.Parse(productPrice) * Int32.Parse(orderVolume) + " рублей");
                Console.WriteLine();
            }
            else Console.WriteLine("По этому товару не было заказов!");
        }

        public void InfoAboutProduct(string path, WorkbookPart wbPart)
        {
            InputNameOfProduct(); //ввод наименования товара
            
            SearchNameAndPrice(path, wbPart); //поиск кода товара и его цены по названию товара по наименованию товара
            SearchCodeAndVolumeAndDate(path, wbPart); //поиск кода клиента, объема и даты заказа по коду товара
            SearchNameOfOrganization(path, wbPart); // поиск названия организации по коду клиента

            OutInfoAboutProduct(); //вывод найденной информации о товаре

            Console.WriteLine("Нажмте любую клавишу чтобы продолжить..");
            Console.ReadKey();
        }
    }
}
