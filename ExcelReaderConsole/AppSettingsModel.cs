using System;

namespace ExcelReaderConsole
{
    [Serializable]
    public class AppSettingsModel
    {
        public string CardFileExtension { get; set; }
        public string ExcelTemplateFilePath { get; set; }
        public string InputDirectoryPath { get; set; }
        public string OutputDirectoryPath { get; set; }
        public string Encoding { get; set; }
    }
}
