using System;
using ExcelOperations.Library;
using Microsoft.Extensions.DependencyInjection;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var serviceProvider = new ServiceCollection()
                .AddSingleton<Spreadsheet>()
                .AddSingleton<ExcelHandler>()
                .BuildServiceProvider();

            var spreadsheet = serviceProvider.GetService<Spreadsheet>();
            var excelHandler = serviceProvider.GetService<ExcelHandler>();

            Console.WriteLine("Enter cell data (format: CellID=Value), type 'done' to finish:");

            while (true)
            {
                Console.Write("Enter cell and value: ");
                string input = Console.ReadLine();

                if (input?.ToLower() == "done")
                {
                    break;
                }

                if (input != null && input.Contains("="))
                {
                    var parts = input.Split('=');
                    if (parts.Length == 2)
                    {
                        string cellId = parts[0].Trim();
                        string cellValue = parts[1].Trim();

                        spreadsheet.SetCell(cellId, cellValue);
                        Console.WriteLine($"Cell {cellId} set to '{cellValue}'");
                    }
                    else
                    {
                        Console.WriteLine("Invalid format. Please use 'CellID=Value'.");
                    }
                }
                else
                {
                    Console.WriteLine("Invalid format. Please use 'CellID=Value'.");
                }
            }

            Console.WriteLine("\nCells with values:");
            foreach (var cellId in spreadsheet.DisplayCellsList())
            {
                var cell = spreadsheet.GetCell(cellId);
                Console.WriteLine($"Cell {cellId}: {cell.GetValue()}");
            }

            Console.Write("Would you like to save to an Excel file? (y/n): ");
            string saveOption = Console.ReadLine();
            if (saveOption?.ToLower() == "y")
            {
                Console.Write("Enter file path to save (e.g., output.xlsx): ");
                string filePath = Console.ReadLine();
                excelHandler.SaveToExcel(filePath, spreadsheet);
                Console.WriteLine($"Spreadsheet saved to {filePath}");
            }

            //spreadsheet.SetCell("A1", "10");
            //spreadsheet.SetCell("A2", "20");
            //Console.WriteLine($"Cell A1: {spreadsheet.GetCell("A1").GetValue()}");
            //Console.WriteLine($"Cell A2: {spreadsheet.GetCell("A2").GetValue()}");

            //Console.WriteLine("Cells with values:");
            //foreach (var cellId in spreadsheet.DisplayCellsList())
            //{
            //    Console.WriteLine(cellId);
            //}

            //Console.WriteLine("Serialized Spreadsheet:");
            //Console.WriteLine(spreadsheet.Serialize());

            //excelHandler.SaveToExcel("output.xlsx", spreadsheet);

            //Console.WriteLine("Excel file saved successfully.");
        }
    }
}