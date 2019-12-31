using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
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

        public void BuildCard(string cardName, Document document)
        {
            string path = Path.Combine(outputDirectoryPath, $"{cardName}{extension}");
            FileInfo cardFileInfo = new FileInfo(path);
            if (!cardFileInfo.Exists)
            {
                using (StreamWriter sw = new StreamWriter(cardFileInfo.OpenWrite(), encoding))
                {
                    foreach(var attribute in document.GetNotNullAttributes())
                    {
                        sw.WriteLine($"{attribute.Key} {attribute.Value.Value}");
                    }
                    sw.Close();
                }
            }
        }

        public void BuildAdditionalCard(string cardName, Document document)
        {
            string path = Path.Combine(outputDirectoryPath, FileManager.Instance.GetAttachDirName(document), $"{cardName}{extension}");
            FileInfo cardFileInfo = new FileInfo(path);
            if (!cardFileInfo.Exists)
            {
                using (StreamWriter sw = new StreamWriter(cardFileInfo.OpenWrite(), encoding))
                {
                    KeyValuePair<string, DocumentAttributeValue> attr7860;
                    try
                    {
                        attr7860 = document.GetNotNullAttributes().First(attribute => attribute.Key.Equals("7860:"));
                    }
                    catch (ArgumentNullException nullEx)
                    {
                        // TODO: notify UI about error in template
                        return;
                    }

                    sw.WriteLine($"1: {attr7860.Value.Value}");
                    sw.WriteLine($"13: {document.CopiedScanFileName}");
                    sw.Close();
                }
            }
        }
    }
}
