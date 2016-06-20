using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace FetchData
{
    /// <summary>
    /// Sample modle for the first three header of the file
    /// </summary>
    public class Sheets
    {
        [JsonProperty("Sheet1")]
        public List<Row> Sheet1 { get; set; }
        [JsonProperty("Sheet2")]
        public List<Row> Sheet2 { get; set; }
    }

    public class Row
    {
        [JsonProperty("one")]
        public string Col1 { get; set; }
        [JsonProperty("two")]
        public string Col2 { get; set; }
        [JsonProperty("three")]
        public string Col3 { get; set; }
    }
}
