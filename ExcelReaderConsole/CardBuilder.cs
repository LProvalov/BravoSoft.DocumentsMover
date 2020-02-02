using System;
using System.IO;
using ExcelReaderConsole.Models;
using System.Text;

namespace ExcelReaderConsole
{
    public class CardBuilder
    {
        public Func<string, bool> FileExsistOverwrite = null;
        private Encoding encoding = Encoding.GetEncoding(AppSettings.Instance.GetEncoding());

        private string extension;
        private string outputDirectoryPath;
        private CardBuilder(string extension)
        {
            this.extension = extension;
            outputDirectoryPath = Directory.GetCurrentDirectory();
        }

        private static CardBuilder instance;
        public static CardBuilder Instance
        {
            get
            {
                if (instance == null)
                {
                    string extension = AppSettings.Instance.GetCardFileExtension();
                    instance = new CardBuilder(extension);
                }
                return instance;
            }
        }

        public void SetOutputDirectory(string path)
        {
            DirectoryInfo outputDirectory = new DirectoryInfo(path);
            if (!outputDirectory.Exists)
            {
                outputDirectory.Create();
            }
            outputDirectoryPath = outputDirectory.FullName;

        }

        public string GetCardPath(string cardName)
        {
            return Path.Combine(outputDirectoryPath, $"{cardName}{extension}");
        }
        public string GetAdditionalCardName(string cardName, Document document)
        {
            return Path.Combine(outputDirectoryPath, FileManager.Instance.GetAttachDirName(document),
                $"{cardName}{extension}");
        }
        public string GetAdditionalCardNameForAttachment(string cardName, Document document)
        {
            return Path.Combine(outputDirectoryPath, FileManager.Instance.GetAttachDirName(document),
                $"{cardName}{extension}");
        }

        private void CreateCard(FileInfo cardFileInfo, Document document)
        {
            using (StreamWriter sw = new StreamWriter(cardFileInfo.OpenWrite(), encoding))
            {
                foreach (var attribute in document.GetNotNullAttributes())
                {
                    sw.WriteLine($"{attribute.Key} {attribute.Value.Value}");
                }
                sw.Close();
            }
        }
        public void BuildCard(string cardName, Document document)
        {
            string path = Path.Combine(outputDirectoryPath, $"{cardName}{extension}");
            FileInfo cardFileInfo = new FileInfo(path);
            if (cardFileInfo.Exists)
            {
                bool ovrewrite = FileExsistOverwrite?.Invoke(cardFileInfo.FullName) ?? false;
                if (ovrewrite)
                {
                    cardFileInfo.Delete();
                    CreateCard(cardFileInfo, document);
                }
            }
            else
            {
                CreateCard(cardFileInfo, document);
            }
        }

        private void CreateAdditionalCard(FileInfo cardFileInfo, Document document)
        {
            using (StreamWriter sw = new StreamWriter(cardFileInfo.OpenWrite(), encoding))
            {
                sw.WriteLine($"1: {document.ScanFileName}");
                sw.WriteLine($"13: {document.CopiedScanFileInfo.Name}");
                sw.Close();
            }
        }
        public void BuildAdditionalCard(string cardName, Document document)
        {
            string path = Path.Combine(outputDirectoryPath, FileManager.Instance.GetAttachDirName(document),
                $"{cardName}{extension}");
            FileInfo cardFileInfo = new FileInfo(path);
            if (cardFileInfo.Exists)
            {
                bool overwrite = FileExsistOverwrite?.Invoke(cardFileInfo.FullName) ?? false;
                if (overwrite)
                {
                    cardFileInfo.Delete();
                    CreateAdditionalCard(cardFileInfo, document);
                }
            }
            else 
            {
                CreateAdditionalCard(cardFileInfo, document);
            }
        }

        private void CreateCardForAttachment(FileInfo cardFileInfo, Document document, int attachmentCount)
        {
            using (StreamWriter sw = new StreamWriter(cardFileInfo.OpenWrite(), encoding))
            {
                sw.WriteLine($"1: {document.AttachmentsFilesInfos[attachmentCount].Name}");
                sw.WriteLine($"13: {document.CopiedAttachmentsFilesInfos[attachmentCount].Name}");
                sw.Close();
            }
        }
        public void BuildAdditionalCardForAttachment(string cardName, int attachmentCount, Document document)
        {
            string path = Path.Combine(outputDirectoryPath, FileManager.Instance.GetAttachDirName(document),
                $"{cardName}{extension}");
            FileInfo cardFileInfo = new FileInfo(path);
            if (cardFileInfo.Exists)
            {
                bool overwrite = FileExsistOverwrite?.Invoke(cardFileInfo.FullName) ?? false;
                if (overwrite)
                {
                    cardFileInfo.Delete();
                    CreateCardForAttachment(cardFileInfo, document, attachmentCount);
                }
            }
            else 
            {
                CreateCardForAttachment(cardFileInfo, document, attachmentCount);
            }
        }
    }
}
