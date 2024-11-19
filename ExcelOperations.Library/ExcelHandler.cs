using System;
using System.IO;

namespace ExcelOperations.Library
{
    public class ExcelHandler
    {
        public void SaveToCsv(string filePath, Spreadsheet data)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            }

            if (!filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                filePath += ".csv";
            }

            using (var writer = new StreamWriter(filePath))
            {
                foreach (var cellId in data.DisplayCellsList())
                {
                    var cellValue = data.EvaluateCell(cellId);
                    writer.WriteLine($"{cellId},{cellValue}");
                }
            }

            Console.WriteLine($"Spreadsheet saved to {filePath}");
        }
    }
}
