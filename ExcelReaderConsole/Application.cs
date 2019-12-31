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

                DocumentStatus documentStatus = CheckingSystem.CheckDocument(document);
                string errorMessage;
                if (documentStatus.TextFileExist)
                {
                    if (!fileManager.TryToCopyTextFile(document, out errorMessage))
                    {
                        documentStatus.TextFileWasMoved = false;
                        documentStatus.StatusMessageAppendLine(errorMessage);
                    }
                    else
                    {
                        documentStatus.TextFileWasMoved = true;
                    }
                }

                if (documentStatus.ScanFileExist)
                {
                    if (!fileManager.TryToCopyScanFile(document, out errorMessage))
                    {
                        documentStatus.StatusMessageAppendLine(errorMessage);
                        documentStatus.ScanFileWasMoved = false;
                    }
                    else
                    {
                        documentStatus.ScanFileWasMoved = true;
                    }
                }

                cardBuilder.BuildCard(document.Identifier, document);
                if (documentStatus.ScanFileExist)
                {
                    cardBuilder.BuildAdditionalCard(document.Identifier, document);
                }
            }
        }
    }
}
