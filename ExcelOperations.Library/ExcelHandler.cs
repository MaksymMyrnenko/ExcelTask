using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOperations.Library
{
    public class ExcelHandler
    {
        public void CreateExcel(string filePath, Spreadsheet data)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                foreach (var cell in data.DisplayCellsList())
                {
                    var cellData = data.GetCell(cell);
                    worksheet.Cell(cell).Value = cellData.GetValue();
                }
                workbook.SaveAs(filePath);
            }
        }

        public void EditExcel(string filePath, Spreadsheet data)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                foreach (var cell in data.DisplayCellsList())
                {
                    var cellData = data.GetCell(cell);
                    worksheet.Cell(cell).Value = cellData.GetValue();
                }
                workbook.Save();
            }
        }

        public Spreadsheet ReadFile(string filePath)
        {
            var spreadsheet = new Spreadsheet();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);

                foreach (var cell in worksheet.CellsUsed())
                {
                    spreadsheet.SetCell(cell.Address.ToString(), cell.GetValue<string>());
                }
            }

            return spreadsheet;
        }

        public void SaveToExcel(string filePath, Spreadsheet data)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            }

            if (!filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                filePath += ".xlsx";
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                foreach (var cell in data.DisplayCellsList())
                {
                    var cellData = data.GetCell(cell);
                    worksheet.Cell(cell).Value = cellData.GetValue();
                }
                workbook.SaveAs(filePath);
            }
        }
    }
}
