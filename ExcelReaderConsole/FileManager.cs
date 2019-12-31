﻿using System;
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

        public void TryToCopyFiles(Document document)
        {
            // TODO: add errors handler
            AppSettings appSettings = AppSettings.Instance;
            DirectoryInfo documentTextDirectory = CreateDirectoriesForTextFiles(document);
            if (!string.IsNullOrEmpty(document.TextFileName))
            {
                string textFileFullPath = Path.Combine(appSettings.GetInputDirectoryPath(), document.TextFileName);
                FileInfo textFileInfo = new FileInfo(textFileFullPath);
                if (textFileInfo.Exists)
                {
                    string newFilePath = Path.Combine(appSettings.GetOutputDirectoryPath(), GetTextDirName(document), document.Identifier + textFileInfo.Extension);
                    File.Copy(textFileInfo.FullName, newFilePath);
                    FileInfo newFileInfo = new FileInfo(newFilePath);
                    if (newFileInfo.Exists)
                    {
                        document.CopiedTextFileName = newFileInfo.Name;
                    }                    
                }
            }

            DirectoryInfo documentAdditionalDirectory = CreateDirectoriesForAdditionalFiles(document);
            if (!string.IsNullOrEmpty(document.ScanFileName))
            {
                string scanFileFullPath = Path.Combine(appSettings.GetInputDirectoryPath(), document.ScanFileName);
                FileInfo scanFileInfo = new FileInfo(scanFileFullPath);
                if (scanFileInfo.Exists)
                {
                    string newFilePath = Path.Combine(appSettings.GetOutputDirectoryPath(), GetAttachDirName(document), scanFileInfo.Name);
                    File.Copy(scanFileInfo.FullName, newFilePath);
                    FileInfo newFileInfo = new FileInfo(newFilePath);
                    if (newFileInfo.Exists)
                    {
                        document.CopiedScanFileName = newFileInfo.Name;
                    }
                }
            }
        }
    }
}
