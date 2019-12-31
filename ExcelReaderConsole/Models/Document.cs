using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
            TextFileName = string.Empty;
            ScanFileName = string.Empty;
            CopiedTextFileName = string.Empty;
            CopiedScanFileName = string.Empty;
        }

        public Document(string identifier, IEnumerable<DocumentAttribute> attributes)
        {
            this.identifier = identifier;
            values = new Dictionary<string, DocumentAttributeValue>(attributes.Count());
            TextFileName = string.Empty;
            ScanFileName = string.Empty;
            CopiedTextFileName = string.Empty;
            CopiedScanFileName = string.Empty;
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

        public string TextFileName { get; set; }

        public string CopiedTextFileName { get; set; }

        public string ScanFileName { get; set; }

        public string CopiedScanFileName { get; set; }

    }
}
