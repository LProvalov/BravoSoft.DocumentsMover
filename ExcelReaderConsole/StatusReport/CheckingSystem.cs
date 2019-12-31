using ExcelReaderConsole.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelReaderConsole.StatusReport
{
    public class CheckingSystem
    {
        private static AppSettings appSettings = AppSettings.Instance;
        private static bool CheckFile(string path, out string statusMessage)
        {
            if (string.IsNullOrEmpty(path))
            {
                statusMessage = $"Path null or empty";
                return false;
            }

            FileInfo fileInfo = new FileInfo(path);
            if (!fileInfo.Exists)
            {
                statusMessage = $"File doesn't exist";
                return false;
            }

            statusMessage = string.Empty;
            return true;
        }

        public static DocumentStatus CheckDocument(Document document)
        {
            DocumentStatus ds = new DocumentStatus(document);
            StringBuilder sb = new StringBuilder();
            string textFileName = document.TextFileName;
            if (string.IsNullOrEmpty(textFileName))
            {
                ds.TextFileExist = false;
                sb.AppendLine($"Text file is null or empty");
            }
            else
            {
                string textFilePath = Path.Combine(appSettings.GetInputDirectoryPath(), textFileName);
                ds.TextFileExist = CheckFile(textFilePath, out var textStatusMessage);
                if (!string.IsNullOrEmpty(textStatusMessage)) sb.AppendLine(textStatusMessage);
            }

            string scanFileName = document.ScanFileName;
            if (string.IsNullOrEmpty(scanFileName))
            {
                ds.ScanFileExist = false;
                sb.AppendLine($"Scan file is null or empty");
            }
            else
            {
                string scanFilePath = Path.Combine(appSettings.GetInputDirectoryPath(), scanFileName);
                ds.ScanFileExist = CheckFile(scanFilePath, out var scanStatusMessage);
                if (!string.IsNullOrEmpty(scanStatusMessage)) sb.AppendLine(scanStatusMessage);
            }

            ds.StatusMessage = sb.ToString();
            return ds;
        }
    }
}
