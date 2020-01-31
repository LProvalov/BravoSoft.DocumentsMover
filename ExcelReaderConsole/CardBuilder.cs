using System.IO;
using ExcelReaderConsole.Models;
using System.Text;

namespace ExcelReaderConsole
{
    public class CardBuilder
    {
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

        public void BuildCard(string cardName, Document document, bool overwrite = false)
        {
            string path = Path.Combine(outputDirectoryPath, $"{cardName}{extension}");
            FileInfo cardFileInfo = new FileInfo(path);
            if (cardFileInfo.Exists && overwrite)
            {
                cardFileInfo.Delete();
                
            }
            else if (cardFileInfo.Exists && !overwrite)
            {
                return;
            }

            using (StreamWriter sw = new StreamWriter(cardFileInfo.OpenWrite(), encoding))
            {
                foreach (var attribute in document.GetNotNullAttributes())
                {
                    sw.WriteLine($"{attribute.Key} {attribute.Value.Value}");
                }
                sw.Close();
            }
        }

        public void BuildAdditionalCard(string cardName, Document document, bool overwrite = false)
        {
            string path = Path.Combine(outputDirectoryPath, FileManager.Instance.GetAttachDirName(document),
                $"{cardName}{extension}");
            FileInfo cardFileInfo = new FileInfo(path);
            if (cardFileInfo.Exists && overwrite)
            {
                cardFileInfo.Delete();
            }
            else if (cardFileInfo.Exists && !overwrite)
            {
                return;
            }
            using (StreamWriter sw = new StreamWriter(cardFileInfo.OpenWrite(), encoding))
            {
                sw.WriteLine($"1: {document.ScanFileName}");
                sw.WriteLine($"13: {document.CopiedScanFileInfo.Name}");
                sw.Close();
            }
        }

        public void BuildAdditionalCardForAttachment(string cardName, int attachmentCount, Document document, bool overwrite = false)
        {
            string path = Path.Combine(outputDirectoryPath, FileManager.Instance.GetAttachDirName(document),
                $"{cardName}{extension}");
            FileInfo cardFileInfo = new FileInfo(path);
            if (cardFileInfo.Exists && overwrite)
            {
                cardFileInfo.Delete();
            }
            else if (cardFileInfo.Exists && !overwrite)
            {
                return;
            }

            using (StreamWriter sw = new StreamWriter(cardFileInfo.OpenWrite(), encoding))
            {
                sw.WriteLine($"1: {document.AttachmentsFilesInfos[attachmentCount].Name}");
                sw.WriteLine($"13: {document.CopiedAttachmentsFilesInfos[attachmentCount].Name}");
                sw.Close();
            }
        }
    }
}
