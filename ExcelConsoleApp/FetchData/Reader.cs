using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace FetchData
{
    /// <summary>
    /// Serializing a DAtaSet from an *.xlsx file
    /// </summary>
    public class Reader
    {
        private string FilePath { get; set; }

        public string SerializeXlsx(string filePath, FileAccess  access = FileAccess.Read, bool isFirstRowHeader = true)
        {
            FileStream fs = null;
            try
            {
                string localPath = new Uri(filePath).LocalPath;
                fs = File.Open(localPath, FileMode.Open, access, FileShare.ReadWrite);
            }
            catch (IOException e)
            {
                Console.WriteLine(e.GetBaseException());
            }

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs);

            excelReader.IsFirstRowAsColumnNames = isFirstRowHeader;

            DataSet excelDataSet = excelReader.AsDataSet();

            string jsonData = DataSetConvertor(excelDataSet);
            return jsonData;
        }

        public string DataSetConvertor(DataSet dataSet)
        {
            var json = JsonConvert.SerializeObject(dataSet);
            return json;
        }

        public Sheets ParseSheet(string jsonData, string sheetName = null)
        {
            return JsonConvert.DeserializeObject<Sheets>(jsonData);
        }
    }
}
