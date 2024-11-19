using System;
using System.Text.RegularExpressions;

namespace ExcelOperations.Library
{
    public static class FormulaEvaluator
    {
        private static readonly Regex cellReferenceRegex = new Regex(@"[A-Z]+[0-9]+", RegexOptions.IgnoreCase);

        public static string EvaluateFormula(string formula, Spreadsheet spreadsheet)
        {
            string evaluatedFormula = formula;

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

            return EvaluateArithmeticExpression(evaluatedFormula).ToString();
        }

        private static float EvaluateArithmeticExpression(string expression)
        {
            var dataTable = new System.Data.DataTable();
            return Convert.ToSingle(dataTable.Compute(expression, ""));
        }
    }
}
