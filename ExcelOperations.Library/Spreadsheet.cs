using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelOperations.Library
{
    public class Spreadsheet
    {
        private readonly List<Cell> cells;

        public Spreadsheet()
        {
            cells = new List<Cell>();
        }

        public void SetCell(string id, string value)
        {
            var cell = cells.FirstOrDefault(c => c.Id == id) ?? new Cell(id);
            cell.RawValue = value;
            cells.Add(cell);
        }

        public Cell GetCell(string id)
        {
            return cells.FirstOrDefault(c => c.Id == id);
        }

        public List<string> DisplayCellsList()
        {
            return cells.Select(c => c.Id).ToList();
        }

        public string EvaluateCell(string cellId)
        {
            var cell = GetCell(cellId);
            if (cell == null)
            {
                return null; // Cell does not exist
            }

            // Check if the cell value is a formula
            if (cell.RawValue.Contains("+") || cell.RawValue.Contains("-") ||
                cell.RawValue.Contains("*") || cell.RawValue.Contains("/"))
            {
                // Evaluate the formula
                return FormulaEvaluator.EvaluateFormula(cell.RawValue, this);
            }

            // Return the raw value for non-formulas
            return cell.RawValue;
        }

    }
}
