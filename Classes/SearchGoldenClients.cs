using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Task3.Classes
{
    internal class SearchGoldenClients
    {
        public int maxNumOfOrders = 0;
        public int indexOfMaxOrderClients = 0;

        public static string ReadCell(string sheetName, string cellName, WorkbookPart wbPart)
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

        public List<InfoAboutClients> FillDatabase(string month, WorkbookPart wbPart)
        {
            string value;
            string value2;
            string date;

            List<InfoAboutClients> clientsList = new List<InfoAboutClients>();
            int lineNumber = 2;
            value = ReadCell("Клиенты", "A" + lineNumber.ToString(), wbPart);
            while (value != null)
            {
                clientsList.Add(new InfoAboutClients(value, ReadCell("Клиенты", "B" + lineNumber.ToString(), wbPart),
                    ReadCell("Клиенты", "C" + lineNumber.ToString(), wbPart), ReadCell("Клиенты", "D" + lineNumber.ToString(), wbPart), 0));

                int i = 2;
                value2 = ReadCell("Заявки", "C" + i.ToString(), wbPart);
                while (value2 != null)
                {
                    date = (new DateTime(1900, 1, 1, 0, 0, 0).AddDays(Int32.Parse(ReadCell("Заявки", "F" + i.ToString(), wbPart)) - 2)).Month.ToString();
                    if (value == value2 && date == month) clientsList[lineNumber - 2].numOfOrders++;
                    
                    i++;
                    value2 = ReadCell("Заявки", "C" + i.ToString(), wbPart);
                }

                if (clientsList[lineNumber - 2].numOfOrders > maxNumOfOrders)
                {
                    maxNumOfOrders = clientsList[lineNumber - 2].numOfOrders;
                    indexOfMaxOrderClients = lineNumber - 2;
                }

                lineNumber++;
                value = ReadCell("Клиенты", "A" + lineNumber.ToString(), wbPart);
            }

            return clientsList;
        }
    
        public void OutGoldenClients(WorkbookPart wbPart)
        {
            Console.Write("Введите номер месяца: ");
            string month = Console.ReadLine();

            List<InfoAboutClients> t = FillDatabase(month, wbPart);

            if (maxNumOfOrders != 0)
                Console.WriteLine("Золотой клиент месяца №" + month + ": "
                + t[indexOfMaxOrderClients].nameOfOrganization + " (Код клиента: " + t[indexOfMaxOrderClients].clientCode
                + "; адрес: " + t[indexOfMaxOrderClients].adressOfOrganizaton + "; контактное лицо: "
                + t[indexOfMaxOrderClients].contactNameOfOrganization + ")");
            else Console.WriteLine("В этом месяце не было закупок!");

            Console.WriteLine("Нажмте любую клавишу чтобы продолжить..");
            Console.ReadKey();
        }
    }
}
