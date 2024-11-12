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

        public string Serialize()
        {
            return Serializer.ToJson(this);
        }
    }
}
