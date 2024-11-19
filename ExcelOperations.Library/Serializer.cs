using System.Text;

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
                string cellId = cells[i];
                string value = data.EvaluateCell(cellId);

                jsonBuilder.Append($"\"{cellId}\": \"{value}\"");

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
