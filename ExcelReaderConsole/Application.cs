using ExcelReaderConsole.Models;
using ExcelReaderConsole.StatusReport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelReaderConsole
{
    public class Application
    {
        private readonly AppSettings appSettings;
        private readonly DocumentsStorage ds;
        private readonly FileManager fileManager;
        private readonly CardBuilder cardBuilder;

        public Application()
        {
            appSettings = AppSettings.Instance;
            appSettings.LoadAppSettings();
            ds = new DocumentsStorage();
            fileManager = FileManager.Instance;
            cardBuilder = CardBuilder.Instance;
        }

        public void Run()
        {
            string excelPath = appSettings.GetExcelTemplateFilePath();
            ExcelReader excelReader = new ExcelReader(excelPath);
            excelReader.ReadData(ds);

            DirectoryInfo outputDirectory = new DirectoryInfo(appSettings.GetOutputDirectoryPath());
            if (!outputDirectory.Exists)
            {
                outputDirectory.Create();
            }
            cardBuilder.SetOutputDirectory(outputDirectory.FullName);

            foreach (var document in ds.GetDocuments())
            {
                Console.WriteLine($"\n{document.Identifier}");
                // Checking of files linked with document
                string errorMessage;
                
                fileManager.TryToCopyTextFile(document, out errorMessage);
                if (!string.IsNullOrEmpty(errorMessage)) Console.WriteLine(errorMessage);
                
                fileManager.TryToCopyScanFile(document, out errorMessage);
                if (!string.IsNullOrEmpty(errorMessage)) Console.WriteLine(errorMessage);

                cardBuilder.BuildCard(document.Identifier, document);
                //if (documentStatus.ScanFileExist)
                //{
                //    cardBuilder.BuildAdditionalCard(document.Identifier, document);
                //}
            }
        }
    }
}
