using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReaderConsole.Models;

namespace GUI.Models
{
    public class DocumentItem
    {
        public DocumentItem(Document document, IEnumerable<DocumentAttribute> attributes)
        {
            Identifier = document.Identifier;
            TextFileName = document.TextFileName;
            ScanFileName = document.ScanFileName;
            TextPdfFileName = document.TextPdfFileName;
            AttachmentsFilesNames = document.AttachmentsFilesNames;

            attributeValues = new Dictionary<string, string>();
            foreach (var attribute in attributes)
            {
                attributeValues.Add(attribute.Identifier, document.GetValue(attribute.Identifier).Value.ToString());
            }

            Warning = false;
        }

        public string Identifier { get; set; }
        public string TextFileName { get; set; }
        public string ScanFileName { get; set; }
        public string TextPdfFileName { get; set; }
        public string AttachmentsFilesNames { get; set; }

        private Dictionary<string, string> attributeValues;
        /*
         * Get attribute value by attributeId
         */
        public string this[string attributeId]
        {
            get
            {
                if (attributeValues.ContainsKey(attributeId))
                {
                    return attributeValues[attributeId];
                }
                return string.Empty;
            }
        }

        public bool Warning { get; set; }

        public bool ProcessedSuccessfully { get; set; }
    }
}
