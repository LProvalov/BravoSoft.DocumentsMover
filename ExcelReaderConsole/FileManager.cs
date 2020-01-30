using System;
using System.IO;
using System.Collections.Generic;

using ExcelReaderConsole.Models;
using ExcelReaderConsole.StatusReport;

namespace ExcelReaderConsole
{
    public class FileManager
    {
        private static FileManager instance;
        private FileManager()
        {

        }
        public static FileManager Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new FileManager();
                }
                return instance;
            }
        }
        private readonly AppSettings appSettings = AppSettings.Instance;
        private DirectoryInfo CreateDir(string dirName)
        {
            string dirPath = Path.Combine(AppSettings.Instance.GetOutputDirectoryPath(), dirName);
            DirectoryInfo dirInfo = new DirectoryInfo(dirPath);
            if (!dirInfo.Exists)
            {
                dirInfo.Create();
            }
            return dirInfo;
        }

        public string GetAttachDirName(Document document)
        {
            return $"{document.Identifier}_attach";
        }

        public string GetTextDirName(Document document)
        {
            return $"{document.Identifier}_text";
        }

        public DirectoryInfo CreateDirectoriesForAdditionalFiles(Document document)
        {
            string dirName = GetAttachDirName(document);
            return CreateDir(dirName);
        }

        public DirectoryInfo CreateDirectoriesForTextFiles(Document document)
        {
            string dirName = GetTextDirName(document);
            return CreateDir(dirName);
        }

        public FileInfo GetScanFile(Document document)
        {
            if (string.IsNullOrEmpty(document.ScanFileName))
            {
                throw new Exception($"{document.Identifier} haven't scan file");
            }
            string scanFilePath = Path.Combine(AppSettings.Instance.GetInputDirectoryPath(), document.ScanFileName);
            return new FileInfo(scanFilePath);
        }

        public FileInfo GetTextFile(Document document)
        {
            if (string.IsNullOrEmpty(document.TextFileName))
            {
                throw new Exception($"{document.Identifier} haven't text file");
            }
            string textFilePath = Path.Combine(AppSettings.Instance.GetInputDirectoryPath(), document.TextFileName);
            return new FileInfo(textFilePath);
        }

        public string MakeNewTextFilePath(Document document)
        {
            return document.TextFileInfo != null ? 
                Path.Combine(appSettings.GetOutputDirectoryPath(),
                GetTextDirName(document),
                $"{document.Identifier}_7845{document.TextFileInfo.Extension}") :
                string.Empty;
        }
        public string MakeNewScanFilePath(Document document)
        {
            return document.ScanFileInfo != null ? 
                Path.Combine(appSettings.GetOutputDirectoryPath(), 
                GetAttachDirName(document), document.ScanFileInfo.Name) : 
                string.Empty;
        }
        public string MakeNewTextPdfPath(Document document)
        {
            return document.TextPdfFileInfo != null ?
                Path.Combine(appSettings.GetOutputDirectoryPath(),
                GetTextDirName(document),
                $"{document.Identifier}{document.TextPdfFileInfo.Extension}") :
                string.Empty;
        }
        public bool NeedToOverwriteAttachments(Document document)
        {
            bool result = false;
            int attachmentCount = 0;
            if (!string.IsNullOrEmpty(document.AttachmentsFilesNames))
            {
                foreach (var fileName in document.AttachmentsFilesNames.Split(';'))
                {
                    string attachmentFileFullPath = Path.Combine(appSettings.GetInputDirectoryPath(), fileName.Trim());
                    FileInfo attachmentFileInfo = new FileInfo(attachmentFileFullPath);
                    if (attachmentFileInfo.Exists)
                    {
                        attachmentCount++;
                        string newAttachmentFileName = Path.Combine(appSettings.GetOutputDirectoryPath(),
                            GetAttachDirName(document),
                            $"Вложение{attachmentCount}{attachmentFileInfo.Extension}");
                        if (File.Exists(newAttachmentFileName))
                        {
                            result = true;
                            break;
                        }
                    }
                }
            }
            return false;
        }

        public bool TryToCopyTextFile(Document document, out string errorMessage, bool overwrite)
        {
            errorMessage = string.Empty;
            DirectoryInfo documentTextDirectory = CreateDirectoriesForTextFiles(document);
            if (!string.IsNullOrEmpty(document.TextFileName))
            {
                if (document.TextFileInfo.Exists)
                {
                    string newFilePath = MakeNewTextFilePath(document);
                    File.Copy(document.TextFileInfo.FullName, newFilePath, overwrite);
                    FileInfo newFileInfo = new FileInfo(newFilePath);
                    if (newFileInfo.Exists)
                    {
                        document.CopiedTextFileInfo = newFileInfo;
                        return true;
                    }
                }
                else
                {
                    errorMessage = $"Text file can't be copied. {document.TextFileInfo.Name} doesn't exist.";
                }
            }
            else
            {
                errorMessage = "Text file of document null or empty. It can't be copied.";
            }
            return false;
        }

        public bool TryToCopyScanFile(Document document, out string errorMessage, bool overwrite)
        {
            errorMessage = string.Empty;
            DirectoryInfo documentAdditionalDirectory = CreateDirectoriesForAdditionalFiles(document);
            if (!string.IsNullOrEmpty(document.ScanFileName))
            {
                string scanFileFullPath = Path.Combine(appSettings.GetInputDirectoryPath(), document.ScanFileName);
                document.ScanFileInfo = new FileInfo(scanFileFullPath);
                if (document.ScanFileInfo.Exists)
                {
                    string newFilePath = MakeNewScanFilePath(document);
                    File.Copy(document.ScanFileInfo.FullName, newFilePath, overwrite);
                    FileInfo newFileInfo = new FileInfo(newFilePath);
                    if (newFileInfo.Exists)
                    {
                        document.CopiedScanFileInfo = newFileInfo;
                        return true;
                    }
                }
                else
                {
                    errorMessage = $"Scan file can't be copied. {scanFileFullPath} doesn't exist.";
                }
            }
            else
            {
                errorMessage = $"Scan file of document null or empty. It can't be copied.";
            }

            return false;
        }

        public bool TryToCopyTextPdfFile(Document document, out string errorMessage, bool overwrite)
        {
            errorMessage = string.Empty;
            DirectoryInfo documentTextDirectory = CreateDirectoriesForTextFiles(document);
            if (!string.IsNullOrEmpty(document.TextPdfFileName))
            {
                string textPdfFileFullPath = Path.Combine(appSettings.GetInputDirectoryPath(), document.TextPdfFileName);
                document.TextPdfFileInfo = new FileInfo(textPdfFileFullPath);
                if (document.TextPdfFileInfo.Exists)
                {
                    string newFilePath = MakeNewTextPdfPath(document);
                    File.Copy(document.TextPdfFileInfo.FullName, newFilePath, overwrite);
                    FileInfo newPdfFileInfo = new FileInfo(newFilePath);
                    if (newPdfFileInfo.Exists)
                    {
                        document.CopiedTextPdfFileInfo = newPdfFileInfo;
                        return true;
                    }
                }
                else
                {
                    errorMessage = $"Text pdf file can't be copied. {textPdfFileFullPath} doesn't exist.";
                }
            }
            else
            {
                errorMessage = "Text pdf file is null or empty. It can't be copied.";
            }
            return false;
        }

        public bool TryToCopyAttachmentFiles(Document document, out string errorMessage, bool overwrite)
        {
            errorMessage = string.Empty;

            if (!string.IsNullOrEmpty(document.AttachmentsFilesNames))
            {
                DirectoryInfo documentAdditionalDirectoryInfo = CreateDirectoriesForAdditionalFiles(document);
                string[] attachmentsFilesNames = document.AttachmentsFilesNames.Split(';');
                List<FileInfo> attachmentFilesInfos = new List<FileInfo>(attachmentsFilesNames.Length);
                List<FileInfo> copiedAttachments = new List<FileInfo>(attachmentsFilesNames.Length);
                int attachmentCount = 0;
                foreach (var fileName in attachmentsFilesNames)
                { 
                    string attachmentFileFullPath = Path.Combine(appSettings.GetInputDirectoryPath(), fileName.Trim());
                    FileInfo attachmentFileInfo = new FileInfo(attachmentFileFullPath);
                    if (attachmentFileInfo.Exists)
                    {
                        attachmentCount++;
                        attachmentFilesInfos.Add(attachmentFileInfo);
                        string newAttachmentFileName = Path.Combine(appSettings.GetOutputDirectoryPath(),
                                GetAttachDirName(document),
                                $"Вложение{attachmentCount}{attachmentFileInfo.Extension}");
                        File.Copy(attachmentFileInfo.FullName, newAttachmentFileName, overwrite);
                        FileInfo newAttachmentFileInfo = new FileInfo(newAttachmentFileName);
                        if (newAttachmentFileInfo.Exists)
                        {
                            copiedAttachments.Add(newAttachmentFileInfo);
                        }
                    }
                    else
                    {
                        errorMessage = $"Attachment file {attachmentFileFullPath} doesn't exsist. It can't be copied.";
                    }
                }
                document.AttachmentsFilesInfos = attachmentFilesInfos.ToArray();
                document.CopiedAttachmentsFilesInfos = copiedAttachments.ToArray();
            }
            else
            {
                errorMessage = "Attachments files is null or empty. They can't be copied.";
            }
            return false;
        }
    }
}
