using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelReaderConsole.Models
{
    public class Document
    {
        private string identifier;
        private Dictionary<string, DocumentAttributeValue> values;

        public Document(IEnumerable<DocumentAttribute> attributes)
        {
            identifier = Utility.GenerateStringIdentifier(Utility.GetSessionSeed());
            values = new Dictionary<string, DocumentAttributeValue>(attributes.Count());
        }

        public Document(string identifier, IEnumerable<DocumentAttribute> attributes)
        {
            this.identifier = identifier;
            values = new Dictionary<string, DocumentAttributeValue>(attributes.Count());
        }

        public string Identifier
        {
            get { return this.identifier; }
        }

        public void SetAttributeValue(string attributeId, DocumentAttributeValue value)
        {
            if (string.IsNullOrEmpty(attributeId) || value == null)
                throw new Exception("AttributeIs or Value can't be null or empty");

            if (!values.ContainsKey(attributeId))
            {
                values.Add(attributeId, value);
            }
            else
            {
                values[attributeId] = value;
            }
        }

        public DocumentAttributeValue GetValue(string attributeId)
        {
            if (string.IsNullOrEmpty(attributeId)) throw new Exception("AttributeId can't be null or empty");

            if (values.ContainsKey(attributeId))
            {
                return values[attributeId];
            }
            else
            {
                return null;
            }
        }

        public string this[string attributeId]
        {
            get
            {
                if (values.ContainsKey(attributeId))
                {
                    return values[attributeId].Value.ToString();
                }

                return string.Empty;
            }
        }

        public IEnumerable<DocumentAttributeValue> GetNotNullValues()
        {
            foreach (var value in values.Values)
            {
                if (value.Value != null)
                {
                    yield return value;
                }
            }
        }

        public IEnumerable<KeyValuePair<string, DocumentAttributeValue>> GetNotNullAttributes()
        {
            foreach (var value in values)
            {
                if (value.Value.Value != null)
                {
                    yield return value;
                }
            }
        }

        public string TextFileName { get; set; } = string.Empty;

        private FileInfo _textFileInfo;
        public FileInfo TextFileInfo
        {
            get
            {
                if (_textFileInfo == null)
                {
                    string textFileFullPath =
                        Path.Combine(AppSettings.Instance.InputDirectoryPath, this.TextFileName);
                    _textFileInfo = new FileInfo(textFileFullPath);
                }
                return _textFileInfo;
            }
            set { _textFileInfo = value; }
        }

        public FileInfo CopiedTextFileInfo { get; set; } = null;

        public string ScanFileName { get; set; } = string.Empty;
        private FileInfo _scanFileInfo;
        public FileInfo ScanFileInfo
        {
            get
            {
                if (_scanFileInfo == null)
                {
                    string scanFileFullPath =
                        Path.Combine(AppSettings.Instance.InputDirectoryPath, this.ScanFileName);
                    _scanFileInfo = new FileInfo(scanFileFullPath);
                }
                return _scanFileInfo;
            }
            set { _scanFileInfo = value; }
        }
        public FileInfo CopiedScanFileInfo { get; set; } = null;

        public string AttachmentsFilesNames { get; set; } = string.Empty;
        public FileInfo[] AttachmentsFilesInfos { get; set; } = null;
        public FileInfo[] CopiedAttachmentsFilesInfos { get; set; } = null;

        public string TextPdfFileName { get; set; } = string.Empty;
        private FileInfo _textPdfFileInfo;
        public FileInfo TextPdfFileInfo
        {
            get
            {
                if (_textPdfFileInfo == null)
                {
                    string textPdfFileFullPath =
                        Path.Combine(AppSettings.Instance.InputDirectoryPath, this.TextPdfFileName);
                    _textPdfFileInfo = new FileInfo(textPdfFileFullPath);
                }
                return _textPdfFileInfo;
            }
            set { _textPdfFileInfo = value; }
        } 

        public FileInfo CopiedTextPdfFileInfo { get; set; } = null;
    }
}
