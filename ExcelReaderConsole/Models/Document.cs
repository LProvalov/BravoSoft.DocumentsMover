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

        public void SetValue(string attributeId, DocumentAttributeValue value)
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
            foreach(var value in values)
            {
                if (value.Value.Value != null)
                {
                    yield return value;
                }
            }
        }

        public string TextFileName { get; set; } = string.Empty;
        public FileInfo TextFileInfo { get; set; } = null;
        public FileInfo CopiedTextFileInfo { get; set; } = null;

        public string ScanFileName { get; set; } = string.Empty;
        public FileInfo ScanFileInfo { get; set; } = null;
        public FileInfo CopiedScanFileInfo { get; set; } = null;

        public string AttachmentsFilesNames { get; set; } = string.Empty;
        public FileInfo[] AttachmentsFilesInfos { get; set; } = null;
        public FileInfo[] CopiedAttachmentsFilesInfos { get; set; } = null;

        public string TextPdfFileName { get; set; } = string.Empty;
        public FileInfo TextPdfFileInfo { get; set; } = null;
        public FileInfo CopiedTextPdfFileInfo { get; set; } = null;
    }
}
