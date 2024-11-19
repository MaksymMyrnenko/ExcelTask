using System;
using System.Text.RegularExpressions;
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

            var cellIdPattern = new Regex(@"^[A-Z]+[0-9]+$", RegexOptions.IgnoreCase);

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

                        if (cellIdPattern.IsMatch(cellId))
                        {
                            spreadsheet.SetCell(cellId, cellValue);
                            Console.WriteLine($"Cell {cellId} set to '{cellValue}'");
                        }
                        else
                        {
                            Console.WriteLine("Invalid cell ID format. Please enter a valid cell (e.g., A1, B2).");
                        }
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

            // Save as CSV for Excel compatibility
            Console.Write("Enter file path to save (e.g., output.csv): ");
            string filePath = Console.ReadLine();
            excelHandler.SaveToCsv(filePath, spreadsheet);
        }
    }
}
