using ExcelReaderConsole.Models;
using System;
using System.IO;

namespace ExcelReaderConsole
{
    public class ConsoleApplication
    {
        private readonly AppSettings appSettings;
        private readonly DocumentsStorage ds;
        private readonly FileManager fileManager;
        private readonly CardBuilder cardBuilder;

        public ConsoleApplication()
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

                fileManager.TryToCopyTextPdfFile(document, out errorMessage);
                if (!string.IsNullOrEmpty(errorMessage)) Console.WriteLine(errorMessage);

                fileManager.TryToCopyAttachmentFiles(document, out errorMessage);
                if (!string.IsNullOrEmpty(errorMessage)) Console.WriteLine(errorMessage);

                cardBuilder.BuildCard(document.Identifier, document);
                if (document.CopiedScanFileInfo != null && document.CopiedScanFileInfo.Exists)
                {
                    cardBuilder.BuildAdditionalCard(document.Identifier, document);
                }

                if (document.CopiedAttachmentsFilesInfos != null && document.CopiedAttachmentsFilesInfos.Length > 0)
                {
                    for (int attachmentCount = 1; attachmentCount <= document.CopiedAttachmentsFilesInfos.Length; attachmentCount++)
                    {
                        cardBuilder.BuildAdditionalCardForAttachment( $"Вложение{attachmentCount}", attachmentCount - 1, document);
                    }
                }
            }
        }
    }
}
