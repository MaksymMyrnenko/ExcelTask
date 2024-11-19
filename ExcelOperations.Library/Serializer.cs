using System.Text;
using System.Collections.Generic;

namespace ExcelOperations.Library
{
    public static class Serializer
    {
        public static string ToJson(Spreadsheet data)
        {
            var jsonBuilder = new StringBuilder();
            jsonBuilder.Append("{");

            var cells = data.DisplayCellsList();
            for (int i = 0; i < cells.Count; i++)
            {
                var cellId = cells[i];
                var cell = data.GetCell(cellId);
                var cellValue = cell.GetValue();

                jsonBuilder.Append($"\"{cellId}\": \"{cellValue}\"");

                // Add a comma after each cell except the last one
                if (i < cells.Count - 1)
                {
                    jsonBuilder.Append(", ");
                }
            }

            jsonBuilder.Append("}");
            return jsonBuilder.ToString();
        }
    }
}
