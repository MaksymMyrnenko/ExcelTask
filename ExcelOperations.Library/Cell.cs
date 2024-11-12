using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public float EvaluateFormula()
        {
            if (float.TryParse(RawValue, out float result))
            {
                return result;
            }
            else
            {
                return 0;
            }
        }
    }
}
