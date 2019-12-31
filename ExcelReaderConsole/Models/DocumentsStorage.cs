using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReaderConsole.Models
{
    public class DocumentsStorage
    {
        private int attributeCount = 0;
        private List<DocumentAttribute> attributes;
        private Dictionary<string, Document> documentsDictionary;

        public DocumentsStorage()
        {
            documentsDictionary = new Dictionary<string, Document>();
        }

        public void Init(int attributeCount)
        {
            this.attributeCount = attributeCount;
            attributes = new List<DocumentAttribute>(attributeCount);
            for (int i = 0; i < attributeCount; i++)
            {
                attributes.Add(new DocumentAttribute());
            }
        }

        public void SetAttributeName(int number, string name)
        {
            if (string.IsNullOrEmpty(name)) throw new Exception("name couldn't be null or empty");
            attributes[number].Name = name;
        }

        public void SetAttributeIdentifier(int number, string identifier)
        {
            if (number > attributeCount) throw new IndexOutOfRangeException($"number is {number}, but it should be less than {attributeCount}");
            if (string.IsNullOrEmpty(identifier)) throw new Exception("idntifier couldn't be null or empty");
            attributes[number].Identifier = identifier;
        }

        public string GetAttributeIdentifier(int number)
        {
            if (number > attributeCount) throw new IndexOutOfRangeException($"number is {number}, but it should be less than {attributeCount}");
            return attributes[number].Identifier;
        }

        public void SetAttributeValue(string documentId, string attributeId, DocumentAttributeValue value)
        {

            if (!documentsDictionary.ContainsKey(documentId)) throw new Exception($"Incorrect documentId: {documentId}");
            if (string.IsNullOrEmpty(attributeId)) throw new Exception("attributeId can't be null or empty");
            if (value == null) throw new Exception("document attribute value can't be null");

            documentsDictionary[documentId].SetValue(attributeId, value);            
        }

        public string AddDocument(string identifier = null)
        {
            Document newDocument = string.IsNullOrEmpty(identifier) ? new Document(attributes) : new Document(identifier, attributes);
            documentsDictionary.Add(newDocument.Identifier, newDocument);
            return newDocument.Identifier;
        }

        public void SetTextFileName(string documentId, string textFileName)
        {
            if (!documentsDictionary.ContainsKey(documentId)) throw new Exception($"Incorrect documentId: {documentId}");
            documentsDictionary[documentId].TextFileName = textFileName;
        }

        public void SetScanFileName(string documentId, string scanFileName)
        {
            if (!documentsDictionary.ContainsKey(documentId)) throw new Exception($"Incorrect documentId: {documentId}");
            documentsDictionary[documentId].ScanFileName = scanFileName;
        }

        public IEnumerable<Document> GetDocuments()
        {
            foreach (var value in documentsDictionary.Values)
            {
                yield return value;
            }
        }

    }
}
