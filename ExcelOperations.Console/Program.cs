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

            while (true)
            {
                Console.WriteLine("\nMain Menu:");
                Console.WriteLine("1) Set values");
                Console.WriteLine("2) Get value from specific cell");
                Console.WriteLine("3) List all values");
                Console.WriteLine("4) Save to file");
                Console.WriteLine("5) Serialize to JSON");
                Console.WriteLine("Enter the number of your choice (or type 'exit' to quit):");

                string choice = Console.ReadLine();
                if (choice?.ToLower() == "exit")
                {
                    break;
                }

                switch (choice)
                {
                    case "1":
                        SetValues(spreadsheet);
                        break;
                    case "2":
                        GetValueFromCell(spreadsheet);
                        break;
                    case "3":
                        ListAllValues(spreadsheet);
                        break;
                    case "4":
                        SaveToFile(excelHandler, spreadsheet);
                        break;
                    case "5":
                        SerializeToJson(spreadsheet);
                        break;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void SetValues(Spreadsheet spreadsheet)
        {
            Console.WriteLine("Enter cell data (format: CellID=Value), type 'done' to finish:");
            var cellIdPattern = new Regex(@"^[A-Z]+[0-9]+$", RegexOptions.IgnoreCase);

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
        }

        private static void GetValueFromCell(Spreadsheet spreadsheet)
        {
            Console.WriteLine("Enter cell IDs to retrieve their values (type 'done' to go back):");

            while (true)
            {
                Console.Write("Enter cell ID: ");
                string cellId = Console.ReadLine();

                if (cellId?.ToLower() == "done")
                {
                    break;
                }

                var value = spreadsheet.EvaluateCell(cellId);
                if (value != null)
                {
                    Console.WriteLine($"Cell {cellId}: {value}");
                }
                else
                {
                    Console.WriteLine($"Cell {cellId} is empty or does not exist.");
                }
            }
        }


        private static void ListAllValues(Spreadsheet spreadsheet)
        {
            Console.WriteLine("Listing all cell values (type 'done' to go back):");

            foreach (var cellId in spreadsheet.DisplayCellsList())
            {
                var evaluatedValue = spreadsheet.EvaluateCell(cellId);
                Console.WriteLine($"Cell {cellId}: {evaluatedValue}");
            }

            Console.WriteLine("Type 'done' to return to the main menu.");
            while (Console.ReadLine()?.ToLower() != "done") { }
        }

        private static void SerializeToJson(Spreadsheet spreadsheet)
        {
            Console.WriteLine("Enter file path to save the JSON (e.g., output.json):");
            string filePath = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("Invalid file path. Serialization aborted.");
                return;
            }

            if (!filePath.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                filePath += ".json";
            }

            try
            {
                string json = Serializer.ToJson(spreadsheet);
                File.WriteAllText(filePath, json);
                Console.WriteLine($"Data successfully serialized to JSON and saved at {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while saving to JSON: {ex.Message}");
            }
        }

        private static void SaveToFile(ExcelHandler excelHandler, Spreadsheet spreadsheet)
        {
            Console.WriteLine("Enter file path to save (e.g., output.csv):");
            string filePath = Console.ReadLine();
            excelHandler.SaveToCsv(filePath, spreadsheet);
        }
    }
}
