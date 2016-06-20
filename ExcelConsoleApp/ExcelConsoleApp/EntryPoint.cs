using System;
using System.Collections.Generic;
using FetchData;
using Newtonsoft.Json;

namespace ExcelConsoleApp
{
    public class EntryPoint
    {
        static void Main(string[] args)
        {
            Reader reader = new Reader();
            // string result = reader.SerializeXlsx(ConfigManager.GetConfigFilePath);
            string result = reader.SerializeXlsx(ConfigManager.GetTestFilePath);

            var sheets = reader.ParseSheet(result);
            var sheet2 = sheets.Sheet2;

            Console.WriteLine("Col 1:\n--------");
            foreach (var root in sheet2)
            {
                Console.WriteLine(root.Col1);
            }

            Console.Read();
        }
    }
}
