namespace ExcelOperations.Library
{
    public class Cell
    {
        public string Id { get; }
        public string RawValue { get; set; }

        public Cell(string id)
        {
            Id = id;
        }

        public string GetValue()
        {
            return RawValue;
        }
    }
}
