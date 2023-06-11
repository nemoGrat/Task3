using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace Task3.Classes
{
    internal class ReWriteClients
    {
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

        public static void UpdateCell(SpreadsheetDocument spreadSheet, string text,  uint rowIndex, string columnName)
        {
            WorksheetPart worksheetPart =  GetWorksheetPartByName(spreadSheet, "Клиенты");

                if (worksheetPart != null)
                {
                    Row row = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
                    Cell cell = row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0).First();

                    cell.CellValue = new CellValue(text);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);

                    worksheetPart.Worksheet.Save();
                    spreadSheet.Save();
                    spreadSheet.Close();
                }
        }

        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document,  string sheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>(). Elements<Sheet>().Where(s => s.Name == sheetName);

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

            return worksheetPart;

        }

        public void NewContactNameOfOrganization(SpreadsheetDocument spreadSheet, WorkbookPart wbPart)
        {
            Console.Write("Введите код клиента для изменения ее контактного лица: ");
            string contactCode = Console.ReadLine();

            uint row = 2;
            string value = ReadCell("Клиенты", "A" + row.ToString(), wbPart);
            bool flag = true;
            while (value != null && flag)
            {
                if (value == contactCode)
                {
                    Console.Write("Введите новое контактное лицо организации \"" + ReadCell("Клиенты", "B" + row.ToString(), wbPart) + "\": ");
                    string newContactName = Console.ReadLine();

                    UpdateCell(spreadSheet, newContactName, row, "D");

                    Console.WriteLine("Контактное лицо изменено!");

                    flag = false;
                }

                row++;
                value = ReadCell("Клиенты", "A" + row.ToString(), wbPart);
            }

            Console.WriteLine("Нажмте любую клавишу чтобы продолжить..");
            Console.ReadKey();
        }

    }
}
