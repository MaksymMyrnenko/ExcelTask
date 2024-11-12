using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOperations.Library
{
    public static class Serializer
    {
        public static string ToJson(Spreadsheet data)
        {
            return JsonConvert.SerializeObject(data);
        }
    }
}
