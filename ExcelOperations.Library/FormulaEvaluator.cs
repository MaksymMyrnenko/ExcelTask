using System;
using System.Text.RegularExpressions;

namespace ExcelOperations.Library
{
    public static class FormulaEvaluator
    {
        private static readonly Regex cellReferenceRegex = new Regex(@"[A-Z]+[0-9]+", RegexOptions.IgnoreCase);

        public static string EvaluateFormula(string formula, Spreadsheet spreadsheet)
        {
            // If the formula contains arithmetic operators, evaluate it
            if (formula.Contains("+") || formula.Contains("-") || formula.Contains("*") || formula.Contains("/"))
            {
                string evaluatedFormula = formula;

                // Replace each cell reference in the formula with its value
                foreach (Match match in cellReferenceRegex.Matches(formula))
                {
                    string cellId = match.Value;
                    string cellValue = spreadsheet.GetCell(cellId)?.GetValue();

                    if (float.TryParse(cellValue, out float numericValue))
                    {
                        evaluatedFormula = evaluatedFormula.Replace(cellId, numericValue.ToString());
                    }
                    else
                    {
                        throw new InvalidOperationException($"Cell {cellId} does not contain a numeric value.");
                    }
                }

                // Evaluate the final arithmetic expression
                return EvaluateArithmeticExpression(evaluatedFormula).ToString();
            }

            // If there are no arithmetic operators, treat as concatenation
            return formula;
        }

        private static float EvaluateArithmeticExpression(string expression)
        {
            // Simple evaluation logic, can be expanded for more operators
            var dataTable = new System.Data.DataTable();
            return Convert.ToSingle(dataTable.Compute(expression, ""));
        }
    }
}
