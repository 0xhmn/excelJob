using System.Configuration;
using System.IO;

namespace ExcelConsoleApp
{
    public static class ConfigManager
    {
        public static string GetConfigFilePath => ConfigurationManager.AppSettings["FilePath"];

        public static string GetTestFilePath
        {
            get
            {
                var rootPath = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)));
                return rootPath != null ? Path.Combine(rootPath, "SampleSheet\\test.xlsx") : null;
            }
        }
    }
}
